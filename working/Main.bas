Attribute VB_Name = "Main"
'Main module
Option Explicit
Public gStep As String

'debug  duplicate pivot tables in country pivot sheet-----------------
Private gBuildRunId As String
Private gBuildCountrySeen As Object
Private gBackupToken As String
Public gUploadId As String
Public gUploadedAt As Date
Dim gDummyPivotCache As PivotCache
Public gRunUnlocked As Boolean
Public gSelectedItems As Variant ' FileDialog.SelectedItems-like collection
Public gChartRefreshBusy As Boolean  ' prevents re-entrant slicer->chart updates
Public gPendingChartCountry As String
Public gLastChartRebuildTime As Double  ' timestamp of last rebuild start (Now)
Public gPendingAxisCountry As String ' country queued for OnTime chart refresh
Public gDeferComboChartName As String   ' reserved (unused in PivotChart strategy)
Public gDeferComboDashName  As String   ' reserved (unused in PivotChart strategy)
' Reserved globals from prior static-chart strategy (kept to avoid compile errors)
Public gRebuildDashName   As String
Public gRebuildChartName  As String
Public gRebuildTitle      As String
Public gRebuildLeft       As Double
Public gRebuildTop        As Double
Public gRebuildWidth      As Double
Public gRebuildHeight     As Double
Public gRebuildBlockAddr  As String
'=========================
' Excel 2019 compatibility constants
'=========================
Private Const XL_SLICER_CROSSFILTER_SHOW_ALL As Long = 0
Private Const XL_SLICER_CROSSFILTER_SHOW_ITEMS_WITH_DATA As Long = 2   ' xlSlicerCrossFilterShowItemsWithData


' Excel 2019-safe enum values (do NOT rely on missing constants)
Private Const XL_SLICER_NO_CROSSFILTER As Long = 1
Private Const XL_SLICER_SHOW_ITEMS_WITH_DATA_AT_TOP As Long = 2
Private Const XL_SLICER_SHOW_ITEMS_WITH_NO_DATA As Long = 3
Private Const XL_SLICER_HIDE_BUTTONS_WITH_NO_DATA As Long = 4   ' <-- what you want



'========================================================
' APAC Treatment Dashboard (Senior Management) - GOLD
'
' Dashboard: KPIs + Slicers + 3 PivotCharts + 2 Breakdown Tables (FORMULA)
' Pivots:     All PivotTables (macro built)
' Calc:       KPI outputs (very hidden)
' Data:       tblData (appended by importer)
'
' One click:
' 1) Select country workbooks (MY/SG/CN/JP/...)
' 2) Import -> Build pivots -> Build slicers/charts -> Render breakdown tables
'
' Expected source workbook format:
' - Each sheet name: "Country - Center" (recommended)
' - Row 4 contains "Day" headers starting each month block
' - Under each block: columns = Day, Budget, Actual, Variance, %
'========================================================
'=========================
' PUBLIC ENTRY POINT
'=========================
'----------------------------------------------------------------
' Returns True if Dashboard charts have already been built.
' Used as smart gate: refresh-only vs full rebuild.
'----------------------------------------------------------------
Private Function DashboardIsBuilt() As Boolean
    DashboardIsBuilt = False
    On Error Resume Next
    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    If wsPvt Is Nothing Then Exit Function
    Dim pt As PivotTable
    Set pt = wsPvt.PivotTables("ptKPI")
    If pt Is Nothing Then Exit Function
    Set pt = wsPvt.PivotTables("ptCountry")
    If pt Is Nothing Then Exit Function
    Set pt = wsPvt.PivotTables("ptMonth")
    If pt Is Nothing Then Exit Function
    Set pt = wsPvt.PivotTables("ptCenter")
    If pt Is Nothing Then Exit Function
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    If wsDash Is Nothing Then Exit Function
    On Error GoTo 0
    DashboardIsBuilt = True
End Function

'----------------------------------------------------------------
' Re-stages tblData to CacheHost with current row count,
' updates SourceData on all pivot caches, then refreshes.
' This ensures new imported rows appear in all pivots.
'----------------------------------------------------------------
Public Sub RefreshAllPivotCaches()
    ' Works for BOTH cases:
    '   - Value-only update (same rows, changed values) -> same row count, values re-copied
    '   - New rows added -> larger row count, SourceData address updated
    ' The pivot cache is always re-staged from current tblData state.
    On Error GoTo SafeExit
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Data").ListObjects("tblData")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    Dim wsHost As Worksheet
    Set wsHost = EnsureCacheHostSheetVisible()
    wsHost.Cells.Clear
    Dim rngSrc As Range
    Set rngSrc = lo.Range
    If lo.ShowTotals Then
        Set rngSrc = rngSrc.Resize(rngSrc.Rows.Count - 1, rngSrc.Columns.Count)
    End If
    Dim rngDst As Range
    Set rngDst = wsHost.Range("A1").Resize(rngSrc.Rows.Count, 12)
    rngDst.Value2 = rngSrc.Value2
    Dim hdr As Variant
    hdr = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                "Budget", "Actual", "Variance", "AchievementPct", _
                "UploadId", "UploadedAt")
    wsHost.Range("A1").Resize(1, 12).Value = hdr
    Dim newSrc As String
    newSrc = wsHost.name & "!" & rngDst.Address(ReferenceStyle:=xlR1C1)

    ' Refresh any existing per-country hosts (zz_CH_<country>) and capture their new source ranges.
    Dim countrySrcMap As Object
    Set countrySrcMap = RestageAllCountryCacheHosts(lo)

    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        On Error Resume Next
        Dim src As String
        src = CStr(pc.SourceData)

        If InStr(1, src, wsHost.name, vbTextCompare) > 0 Then
            pc.SourceData = newSrc
        Else
            Dim hostNm As Variant
            For Each hostNm In countrySrcMap.keys
                If InStr(1, src, CStr(hostNm), vbTextCompare) > 0 Then
                    pc.SourceData = CStr(countrySrcMap(hostNm))
                    Exit For
                End If
            Next hostNm
        End If
        pc.Refresh
        Err.Clear
        On Error GoTo SafeExit
    Next pc

    ' Hard-refresh main dashboard pivots from the newly staged source.
    ' This avoids stale values when cache SourceData string parsing fails.
    RefreshMainPivotsAgainstSource newSrc

    wsHost.Visible = xlSheetVeryHidden
SafeExit:
End Sub

Private Sub RefreshMainPivotsAgainstSource(ByVal srcR1C1 As String)
    On Error GoTo SafeExit
    If Len(srcR1C1) = 0 Then Exit Sub

    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    If wsPvt Is Nothing Then Exit Sub

    Dim names As Variant
    names = Array("ptKPI", "ptCountry", "ptMonth", "ptCenter", "ptCountryMonth", "ptCountryCenter")

    ' 1) Collect unique cache indexes currently used by main pivots.
    Dim cacheIdx As Object
    Set cacheIdx = CreateObject("Scripting.Dictionary")
    cacheIdx.CompareMode = vbTextCompare

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim pt As PivotTable
        On Error Resume Next
        Set pt = wsPvt.PivotTables(CStr(names(i)))
        On Error GoTo SafeExit
        If Not pt Is Nothing Then
            cacheIdx(CStr(pt.PivotCache.Index)) = True
        End If
    Next i

    ' 2) Update those exact caches to the newly staged source.
    Dim k As Variant
    For Each k In cacheIdx.keys
        On Error Resume Next
        Dim pc As PivotCache
        Set pc = ThisWorkbook.PivotCaches(CLng(k))
        If Not pc Is Nothing Then
            pc.SourceData = srcR1C1
            pc.Refresh
        End If
        Err.Clear
        On Error GoTo SafeExit
    Next k

    ' 3) Refresh main pivots explicitly.
    For i = LBound(names) To UBound(names)
        Dim pt2 As PivotTable
        On Error Resume Next
        Set pt2 = wsPvt.PivotTables(CStr(names(i)))
        If Not pt2 Is Nothing Then
            pt2.ManualUpdate = False
            pt2.RefreshTable
        End If
        Err.Clear
        On Error GoTo SafeExit
    Next i

SafeExit:
End Sub

Private Function GetTblDataPivotSourceRows() As Long
    On Error GoTo SafeExit
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Data").ListObjects("tblData")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetTblDataPivotSourceRows = lo.DataBodyRange.Rows.Count + 1  ' + header row
SafeExit:
End Function

Private Function ExtractSourceLastRow(ByVal src As String) As Long
    On Error GoTo SafeExit
    Dim p As Long
    Dim i As Long
    Dim ch As String
    Dim digits As String

    ' R1C1 style, e.g. zz_CacheHost!R1C1:R874C12
    p = InStrRev(src, ":R")
    If p > 0 Then
        digits = ""
        For i = p + 2 To Len(src)
            ch = Mid$(src, i, 1)
            If ch Like "#" Then
                digits = digits & ch
            Else
                Exit For
            End If
        Next i
        If Len(digits) > 0 Then
            ExtractSourceLastRow = CLng(digits)
            Exit Function
        End If
    End If

    ' A1 style fallback, e.g. zz_CacheHost!$A$1:$L$874
    p = InStrRev(src, "$")
    If p > 0 Then
        digits = ""
        For i = p + 1 To Len(src)
            ch = Mid$(src, i, 1)
            If ch Like "#" Then
                digits = digits & ch
            Else
                Exit For
            End If
        Next i
        If Len(digits) > 0 Then
            ExtractSourceLastRow = CLng(digits)
        End If
    End If
SafeExit:
End Function

Public Function MainPivotSourceIsCurrent() As Boolean
    On Error GoTo SafeExit
    Dim expectedRows As Long
    expectedRows = GetTblDataPivotSourceRows()
    If expectedRows <= 1 Then Exit Function

    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    If wsPvt Is Nothing Then Exit Function

    Dim ptM As PivotTable, ptC As PivotTable
    Set ptM = wsPvt.PivotTables("ptMonth")
    Set ptC = wsPvt.PivotTables("ptCountry")
    If ptM Is Nothing Or ptC Is Nothing Then Exit Function

    Dim mRows As Long, cRows As Long
    mRows = ExtractSourceLastRow(CStr(ptM.PivotCache.SourceData))
    cRows = ExtractSourceLastRow(CStr(ptC.PivotCache.SourceData))

    MainPivotSourceIsCurrent = (mRows = expectedRows And cRows = expectedRows)
SafeExit:
End Function

Public Sub ClearMainDashboardSlicerFilters()
    On Error GoTo SafeExit
    On Error Resume Next
    ThisWorkbook.SlicerCaches("scCountry").ClearManualFilter
    ThisWorkbook.SlicerCaches("scMonth").ClearManualFilter
    Err.Clear
    On Error GoTo SafeExit

    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    If wsPvt Is Nothing Then Exit Sub

    Dim pt As PivotTable
    On Error Resume Next
    Set pt = wsPvt.PivotTables("ptMonth")
    If Not pt Is Nothing Then
        pt.ManualUpdate = False
        pt.RefreshTable
    End If
    Set pt = wsPvt.PivotTables("ptCountry")
    If Not pt Is Nothing Then
        pt.ManualUpdate = False
        pt.RefreshTable
    End If
    On Error GoTo 0

SafeExit:
End Sub

Private Function RestageAllCountryCacheHosts(ByVal lo As ListObject) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase$(left$(ws.name, 6)) = "zz_ch_" Then
            Dim countryKey As String
            countryKey = NormalizeCountryKey(Mid$(ws.name, 7))

            Dim src As String
            src = RestageOneCountryCacheHost(lo, countryKey, ws)
            If Len(src) > 0 Then
                map(ws.name) = src
            End If
        End If
    Next ws

    Set RestageAllCountryCacheHosts = map
End Function

Private Function RestageOneCountryCacheHost(ByVal lo As ListObject, ByVal countryKey As String, ByVal wsHost As Worksheet) As String
    On Error GoTo SafeExit
    wsHost.Visible = xlSheetVisible
    modSecurity.UnprotectSheet wsHost

    wsHost.Cells.Clear
    wsHost.Cells.ClearFormats
    wsHost.Cells.UnMerge

    Dim hdr As Variant
    hdr = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                "Budget", "Actual", "Variance", "AchievementPct", _
                "UploadId", "UploadedAt")
    wsHost.Range("A1").Resize(1, 12).Value = hdr

    Dim outRow As Long
    outRow = 2

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim rowKey As String
        Dim regionVal As String
        Dim centerVal As String
        Dim monthVal As String

        rowKey = NormalizeCountryKey(lo.DataBodyRange.Cells(r, 1).Value2)
        regionVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        centerVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2))
        monthVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))

        If rowKey = countryKey _
           And Len(regionVal) > 0 _
           And Len(centerVal) > 0 _
           And Len(monthVal) > 0 Then
            wsHost.Cells(outRow, 1).Resize(1, 12).Value2 = lo.DataBodyRange.Cells(r, 1).Resize(1, 12).Value2
            wsHost.Cells(outRow, 1).Value2 = countryKey
            outRow = outRow + 1
        End If
    Next r

    Dim dataRows As Long
    dataRows = outRow - 2
    If dataRows <= 0 Then GoTo SafeExit

    Dim rngDst2 As Range
    Set rngDst2 = wsHost.Range("A1").Resize(1 + dataRows, 12)
    ValidatePivotSourceRange rngDst2

    RestageOneCountryCacheHost = wsHost.name & "!" & rngDst2.Address(ReferenceStyle:=xlR1C1)

SafeExit:
    On Error Resume Next
    wsHost.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Function

'----------------------------------------------------------------
' Refresh path for subsequent imports and slicer changes.
' Updates static ranges (AA1/AH1) with correct Ach%=SUM/SUM,
' then calls SetSourceData on each chart - series/combo preserved.
'----------------------------------------------------------------

Public Sub RepositionDashboardCharts()
    ' Always enforces correct position/size for chMonth and chCountry.
    ' Called from both build path and refresh path so charts are always correct.
    On Error GoTo SafeExit
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo SafeExit
    If wsDash Is Nothing Then Exit Sub

    Const CHART_MARGIN As Double = 6
    Const CHART_H As Double = 280

    ' Chart top = bottom of lowest slicer + margin
    Dim chTop As Double
    chTop = 0
    Dim shpScan As Shape
    For Each shpScan In wsDash.Shapes
        If shpScan.Type = 24 Then   ' msoSlicer
            If shpScan.top + shpScan.height > chTop Then
                chTop = shpScan.top + shpScan.height
            End If
        End If
    Next shpScan
    If chTop = 0 Then chTop = wsDash.Range("A13").top
    chTop = chTop + CHART_MARGIN

    ' Total usable width from col A left edge to rightmost slicer/button right edge
    Dim totalW As Double
    totalW = 0
    For Each shpScan In wsDash.Shapes
        If shpScan.left + shpScan.width > totalW Then
            totalW = shpScan.left + shpScan.width
        End If
    Next shpScan
    If totalW < 600 Then totalW = wsDash.Range("A1:V1").width

    Dim chW As Double
    Dim chLeft As Double, chRight As Double
    chLeft = wsDash.Range("A1").left + CHART_MARGIN
    chW = (totalW / 2) - CHART_MARGIN * 1.5
    chRight = chLeft + chW + CHART_MARGIN

    ' Apply to chMonth
    On Error Resume Next
    Dim shpM As Shape
    Set shpM = wsDash.Shapes("chMonth")
    If Not shpM Is Nothing Then
        shpM.left = chLeft
        shpM.top = chTop
        shpM.width = chW
        shpM.height = CHART_H
    End If

    ' Apply to chCountry
    Dim shpC As Shape
    Set shpC = wsDash.Shapes("chCountry")
    If Not shpC Is Nothing Then
        shpC.left = chRight
        shpC.top = chTop
        shpC.width = chW
        shpC.height = CHART_H
    End If
    On Error GoTo SafeExit
SafeExit:
End Sub
Public Sub RefreshDashboardCharts(Optional ByVal forceRebind As Boolean = False)
    On Error GoTo SafeExit
    Dim wsPvt As Worksheet, wsDash As Worksheet
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo SafeExit
    If wsPvt Is Nothing Or wsDash Is Nothing Then Exit Sub

    Dim ptMonth As PivotTable, ptCountry As PivotTable
    On Error Resume Next
    Set ptMonth = wsPvt.PivotTables("ptMonth")
    Set ptCountry = wsPvt.PivotTables("ptCountry")
    On Error GoTo SafeExit

    If forceRebind Then
        If Not ptMonth Is Nothing Then
            If Not HasAchDataField(ptMonth) Then EnsureAchievementDataField ptMonth
        End If
        If Not ptCountry Is Nothing Then
            If Not HasAchDataField(ptCountry) Then EnsureAchievementDataField ptCountry
        End If

        If Not ptMonth Is Nothing Then
            ptMonth.ManualUpdate = False
            ptMonth.RefreshTable
        End If

        If Not ptCountry Is Nothing Then
            ptCountry.ManualUpdate = False
            ptCountry.RefreshTable
        End If
    End If

    EnsureMainDashboardPivotCharts wsDash, wsPvt, ptMonth, ptCountry, forceRebind
SafeExit:
End Sub

Private Function HasUsableChartShape(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Function
    Dim shp As Shape
    Set shp = ws.Shapes(shapeName)
    If shp Is Nothing Then Exit Function
    HasUsableChartShape = shp.HasChart
SafeExit:
End Function

Private Sub EnsureMainDashboardPivotCharts( _
    ByVal wsDash As Worksheet, ByVal wsPvt As Worksheet, _
    ByVal ptMonth As PivotTable, ByVal ptCountry As PivotTable, _
    Optional ByVal forceRebuild As Boolean = False)
    On Error GoTo SafeExit
    If wsDash Is Nothing Or wsPvt Is Nothing Then Exit Sub

    Dim rebuildMonth As Boolean
    rebuildMonth = forceRebuild Or (Not HasUsableChartShape(wsDash, "chMonth"))
    If Not ptMonth Is Nothing Then
        If rebuildMonth Then
            Dim shpMonth As Shape
            Set shpMonth = AddPivotChartClean(wsDash, wsPvt, ptMonth, xlColumnClustered, _
                "Monthly Budget vs Actual (Overall)", 20, 200, 400, 280, "chMonth")
        Else
            On Error Resume Next
            ApplyPivotChartComboFormat wsDash.Shapes("chMonth").Chart
            On Error GoTo SafeExit
        End If
    End If

    Dim rebuildCountry As Boolean
    rebuildCountry = forceRebuild Or (Not HasUsableChartShape(wsDash, "chCountry"))
    If Not ptCountry Is Nothing Then
        If rebuildCountry Then
            Dim shpCountry As Shape
            Set shpCountry = AddPivotChartClean(wsDash, wsPvt, ptCountry, xlColumnClustered, _
                "Country Achievement (Overall)", 420, 200, 400, 280, "chCountry")
        Else
            On Error Resume Next
            ApplyPivotChartComboFormat wsDash.Shapes("chCountry").Chart
            On Error GoTo SafeExit
        End If
    End If

SafeExit:
End Sub

Public Sub UpsertMainDashboardComboChart( _
    ByVal wsDash As Worksheet, ByVal wsPvt As Worksheet, _
    ByVal chartName As String, ByVal startAddr As String, _
    ByVal chartTitle As String, _
    ByVal chartLeft As Double, ByVal chartTop As Double, _
    ByVal chartWidth As Double, ByVal chartHeight As Double)

    On Error GoTo SafeExit
    If wsDash Is Nothing Or wsPvt Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = wsPvt.Range(startAddr)

    ' Count data rows in first column of static block (header at row 1).
    Dim nRows As Long
    Dim r As Long
    For r = 1 To 199
        If Len(Trim$(CStr(anchor.Offset(r, 0).Value))) = 0 Then Exit For
        nRows = r
    Next r
    If nRows < 1 Then Exit Sub

    Dim srcRng As Range
    Set srcRng = anchor.Resize(nRows + 1, 4)

    Dim shp As Shape
    Dim ch As Chart
    On Error Resume Next
    Set shp = wsDash.Shapes(chartName)
    On Error GoTo SafeExit

    If shp Is Nothing Then
        On Error Resume Next
        Set shp = wsDash.Shapes.AddChart(xlColumnClustered, chartLeft, chartTop, chartWidth, chartHeight)
        shp.name = chartName
        On Error GoTo SafeExit
    ElseIf Not shp.HasChart Then
        On Error Resume Next
        shp.Delete
        On Error GoTo SafeExit
        Set shp = wsDash.Shapes.AddChart(xlColumnClustered, chartLeft, chartTop, chartWidth, chartHeight)
        shp.name = chartName
    End If

    Set ch = shp.Chart
    On Error Resume Next
    ch.PlotVisibleOnly = False
    On Error GoTo SafeExit
    ' Keep chart bound to column-wise static range so slicer filters do not flip orientation.
    ch.SetSourceData Source:=srcRng, PlotBy:=xlColumns
    ch.HasTitle = True
    ch.chartTitle.Text = chartTitle
    ch.HasLegend = True

    ApplyComboToChart ch

SafeExit:
End Sub

Public Sub ImportAndRefresh()

    gStep = "ImportAndRefresh: begin"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ProgressStart "APAC Dashboard Import", 100
    ProgressSet 1, "Preparing..."

    '--------------------------------------------
    ' 1) Pick + validate files (NO unlock yet)
    '--------------------------------------------
    gStep = "ImportAndRefresh: PickAndValidateFiles"
    If Not PickAndValidateFiles(gSelectedItems) Then
        ProgressFinish "Import cancelled."
        GoTo CleanExit_Cancel
    End If

    '--------------------------------------------
    ' 2) User confirmed selection -> unlock ONCE
    '--------------------------------------------
    gStep = "ImportAndRefresh: RunUnlockAll"
    RunUnlockAll

    ' Init upload run AFTER confirmation
    gStep = "ImportAndRefresh: init upload run"
    gUploadedAt = Now
    gUploadId = BuildUploadId(gUploadedAt)

    '--------------------------------------------
    ' 3) Import data now (assumes unprotected)
    '--------------------------------------------
    gStep = "ImportAndRefresh: ImportSelectedFilesToTblData"
    ImportSelectedFilesToTblData gSelectedItems

    '--------------------------------------------
    ' 4) Your existing build pipeline (NO protect calls)
    '--------------------------------------------
    gStep = "ImportAndRefresh: Backup BEFORE"
    gBackupToken = modBackup.backupToken()
    modBackup.BackupWorkbookCopy "Before", gBackupToken

    If DashboardIsBuilt() Then
        ' --- REFRESH PATH (subsequent imports) ---
        ' Pivots + charts already exist. Just re-stage tblData,
        ' refresh pivot caches, update static ranges, update charts.
        ' No delete, no recreate - combo layout is permanent.
        gStep = "ImportAndRefresh: refresh pivot caches"
        ProgressSet 70, "Refreshing pivots..."
        RefreshAllPivotCaches

        If Not MainPivotSourceIsCurrent() Then
            gStep = "ImportAndRefresh: self-heal main pivots"
            ProgressSet 78, "Rebuilding main pivots..."
            EnsurePivotsSlicersChartsAndTables
        End If

        gStep = "ImportAndRefresh: clear main slicers"
        ClearMainDashboardSlicerFilters

        gStep = "ImportAndRefresh: refresh static ranges + charts"
        ProgressSet 85, "Refreshing charts..."
        RefreshDashboardCharts True
        RepositionDashboardCharts
        EnsureCountryDashboardCharts

        gStep = "ImportAndRefresh: BuildUploadSummary"
        BuildUploadSummary
    Else
        ' --- BUILD PATH (first run) ---
        gStep = "ImportAndRefresh: BuildCountryDashboards"
        ProgressSet 70, "Refreshing dashboards..."
        BuildCountryDashboards

        gStep = "ImportAndRefresh: EnsurePivotsSlicersChartsAndTables"
        ProgressSet 85, "Building pivots / slicers / charts..."
        EnsurePivotsSlicersChartsAndTables
        RepositionDashboardCharts
        EnsureCountryDashboardCharts

        gStep = "ImportAndRefresh: BuildUploadSummary"
        BuildUploadSummary
    End If

    gStep = "ImportAndRefresh: BuildLatestUpdatePivot"
    BuildLatestUpdatePivot

    gStep = "ImportAndRefresh: InjectPivotSheetEvents"
    InjectPivotSheetEvents

    gStep = "ImportAndRefresh: EnforceSheetPositions"
    ProgressSet 90, "Finalizing..."
    EnforceSheetPositions

    ' Land on Dashboard
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
    On Error GoTo CleanFail

    ' Save
    gStep = "ImportAndRefresh: Save workbook (commit)"
    ProgressSet 92, "Saving workbook..."
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    ' Export PDF (optional)
    ProgressSet 95, "Exporting charts to PDF..."
    ExportAllCharts_Fast_ChartSheets_Direct

    gStep = "ImportAndRefresh: Backup AFTER"
    modBackup.BackupWorkbookCopy "After", gBackupToken

    ProgressFinish "Completed"

CleanExit_Success:
    On Error Resume Next
    ProgressForceClose

    ' Always reharden at END (success)
    RunReharden

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
    Exit Sub

CleanExit_Cancel:
    On Error Resume Next
    ProgressForceClose

    ' Cancel = do NOT unlock / do NOT change protection
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
    Exit Sub

CleanFail:
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description

    On Error Resume Next
    ProgressForceClose

    ' If we unlocked, reharden immediately
    If gRunUnlocked Then RunReharden

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "ImportAndRefresh failed at: " & gStep & vbCrLf & _
           errNo & " - " & errMsg, vbCritical
End Sub



Private Sub ImportOneWorkbook(ByVal filePath As String, ByVal lo As ListObject, ByVal keyIndex As Object)
    gStep = "ImportOneWorkbook: open " & filePath

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    gStep = "ImportOneWorkbook: iterate worksheets"
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        gStep = "ImportOneWorkbook: ImportOneSheet [" & wb.name & "!" & ws.name & "]"
        ImportOneSheet ws, lo, keyIndex
    Next ws

    gStep = "ImportOneWorkbook: close " & wb.name
    wb.Close SaveChanges:=False

    gStep = "ImportOneWorkbook: done"
End Sub



Private Sub ImportOneSheet(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal keyIndex As Object)
    gStep = "ImportOneSheet: start [" & ws.Parent.name & "!" & ws.name & "]"

    Dim country As String, region As String, center As String
    gStep = "ImportOneSheet: ParseCountryCenter"
    ParseCountryRegionCenter ws.name, country, region, center

    Dim lastCol As Long
    gStep = "ImportOneSheet: detect lastCol (row 4)"
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    Dim dayCols() As Long, n As Long
    n = 0

    gStep = "ImportOneSheet: scan day headers"
    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String
        hdr = LCase$(Trim$(CStr(ws.Cells(4, c).Value)))
        If InStr(1, hdr, "day") > 0 Then
            n = n + 1
            ReDim Preserve dayCols(1 To n)
            dayCols(n) = c
        End If
    Next c

    If n = 0 Then
        gStep = "ImportOneSheet: no day headers, exit [" & ws.Parent.name & "!" & ws.name & "]"
        Exit Sub
    End If

    Dim lastRow As Long
    gStep = "ImportOneSheet: detect lastRow using first day col"
    lastRow = ws.Cells(ws.Rows.Count, dayCols(1)).End(xlUp).Row
    If lastRow < 5 Then
        gStep = "ImportOneSheet: lastRow < 5, exit [" & ws.Parent.name & "!" & ws.name & "]"
        Exit Sub
    End If

    gStep = "ImportOneSheet: iterate data rows"
    Dim r As Long, k As Long
    For r = 5 To lastRow
        For k = 1 To n
            Dim dc As Long, bc As Long, ac As Long
            dc = dayCols(k)
            bc = dc + 1
            ac = dc + 2

            Dim vDay As Variant, vBudget As Variant, vActual As Variant
            vDay = ws.Cells(r, dc).Value
            If Not IsDate(vDay) Then GoTo NextBlock

            vBudget = ws.Cells(r, bc).Value
            vActual = ws.Cells(r, ac).Value
            If (IsEmpty(vBudget) And IsEmpty(vActual)) Then GoTo NextBlock

            Dim d As Date
            d = CDate(vDay)

            Dim budget As Double, actual As Double
            budget = NzD(vBudget)
            actual = NzD(vActual)

            Dim variance As Double
            variance = actual - budget

            Dim ach As Variant
            If budget = 0 Then
                ach = Empty
            Else
                ach = actual / budget
            End If


            gStep = "ImportOneSheet: UpsertRow [" & country & " | " & center & " | " & Format$(d, "yyyy-mm-dd") & "]"
            Dim mStart As Date
            mStart = DateSerial(Year(d), Month(d), 1)
            UpsertRow lo, keyIndex, country, region, center, d, mStart, _
                Format$(mStart, "mmm-yyyy"), budget, actual, variance, ach



NextBlock:
        Next k
    Next r

    gStep = "ImportOneSheet: done [" & ws.Parent.name & "!" & ws.name & "]"
End Sub




Private Sub ParseCountryRegionCenter(ByVal sheetName As String, _
                                     ByRef country As String, _
                                     ByRef region As String, _
                                     ByRef center As String)

    gStep = "ParseCountryRegionCenter: [" & sheetName & "]"

    Dim parts() As String
    parts = Split(sheetName, "-")

    ' Must be EXACTLY 3 segments
    If UBound(parts) <> 2 Then
        Err.Raise vbObjectError + 701, , _
            "Invalid sheet name format: [" & sheetName & "]" & vbCrLf & _
            "Required format: Country - Region - Center" & vbCrLf & _
            "Example: MY-R1-Landmark"
    End If

    country = Trim$(parts(0))
    region = Trim$(parts(1))
    center = Trim$(parts(2))

    ' Disallow empty segments
    If Len(country) = 0 Or Len(region) = 0 Or Len(center) = 0 Then
        Err.Raise vbObjectError + 702, , _
            "Sheet name contains empty segment: [" & sheetName & "]" & vbCrLf & _
            "Required format: Country - Region - Center"
    End If

    ' Optional: normalize casing
    country = UCase$(country)
    region = UCase$(region)       ' R1, R2, R3...
    center = Trim$(center)

End Sub



Private Function NzD(ByVal v As Variant) As Double
    gStep = "NzD"

    If IsError(v) Then
        NzD = 0
    ElseIf IsEmpty(v) Or v = "" Then
        NzD = 0
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0
    End If
End Function


'=========================
' GOLD BUILDER:
' - pivots on Pivots sheet
' - slicers/charts on Dashboard
' - breakdown tables on Dashboard via formulas
'=========================
Public Sub EnsurePivotsSlicersChartsAndTables()
    gStep = "EnsurePivotsSlicersChartsAndTables: start"
    On Error GoTo EH
    
    ' --- HARD RESET (no PivotTable conflicts) ---
    ' Keep Data and Dashboard; recreate Pivots and Calc from scratch
    Dim wsData As Worksheet, wsDash As Worksheet
    gStep = "EnsurePivotsSlicersChartsAndTables: Set wsData/wsDash"
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Unprotect Dashboard"
    On Error Resume Next
    'wsDash.Unprotect Password:=SYS_SHEET_PWD
    Err.Clear
    On Error GoTo EH
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Delete slicer caches"
    ' Delete slicer caches first (they can lock shapes)
    DeleteSlicerCacheIfExists "scCountry"
    DeleteSlicerCacheIfExists "scMonth"
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Clear dashboard objects"
    ' Clear dashboard visuals (charts/slicers/buttons)
    ClearDashboardObjects wsDash
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Clear dashboard formula areas"
    ' Clear dashboard formula table areas only (safe)
    wsDash.Range("A45:Z120").ClearContents
    
    Dim wsPvt As Worksheet, wsCalc As Worksheet
    gStep = "EnsurePivotsSlicersChartsAndTables: Set wsPvt/wsCalc"
    Set wsPvt = RecreateSheet("Pivots", wsDash)
    
    gStep = "EnsurePivotsSlicersChartsAndTables: EnforceSheetPositions"
    EnforceSheetPositions
            
    ' --- Reset Calc (workbook structure must be unprotected to change Visible) ---
    gStep = "EnsurePivotsSlicersChartsAndTables: Reset Calc"
    
    '==================================================
    ' Reset Calc (must unlock structure to change Visible)
    '==================================================
    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock("Reset Calc")
    
    Dim errNo As Long, errMsg As String
    
    On Error GoTo EH_Calc
    
    Set wsCalc = EnsureCalcSheet()

    
    ' If Calc sheet is protected, clearing will fail
    On Error Resume Next
    'wsCalc.Unprotect Password:=SYS_SHEET_PWD
    On Error GoTo EH_Calc
    
    wsCalc.Visible = xlSheetVisible
    wsCalc.Cells.ClearContents
    wsCalc.Cells.ClearFormats
    wsCalc.Visible = xlSheetVeryHidden
    
    GoTo Clean_Calc
    
EH_Calc:
        ' ? capture first, BEFORE any cleanup that may clear Err
        errNo = Err.Number
        errMsg = Err.Description
    
        'EndStructureUnlock wasProtected
        Err.Raise vbObjectError + 991, , "Reset Calc failed: " & errNo & " - " & errMsg
    
Clean_Calc:
        'EndStructureUnlock wasProtected
        On Error GoTo EH

    '--------------------------------------------------
    ' Validate data table
    '--------------------------------------------------
    gStep = "EnsurePivotsSlicersChartsAndTables: Validate tblData"
    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")
    
    If lo Is Nothing Then
        MsgBox "tblData not found on Data sheet.", vbCritical
        Exit Sub
    End If
    
    If lo.DataBodyRange Is Nothing Then
        MsgBox "No data available in tblData.", vbExclamation
        Exit Sub
    End If

    '--------------------------------------------------
    ' PivotCache (create ONCE, using sheet-local string)
    '--------------------------------------------------
    gStep = "EnsurePivotsSlicersChartsAndTables: Create PivotCache"
    
    ' Pivot cache creation can fail if source sheet is protected
    On Error Resume Next
    'wsData.Unprotect Password:=SYS_SHEET_PWD
    On Error GoTo EH
    
    Dim pc As PivotCache
    
'Debug.Print "EnsurePivotsSlicersChartsAndTables PivotCaches.Count BEFORE create =", ThisWorkbook.PivotCaches.Count


    Set pc = CreatePivotCacheFromTable(lo)    ' <--- uses 'Data'!R1C1:RxxxxCxx
    If pc Is Nothing Then Err.Raise vbObjectError + 882, , "CreatePivotCacheFromTable returned Nothing."
    


    '==================================================
    ' Pivot 1: KPI (Pivots!A1)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptKPI"
    
    ' --- hard guarantee destination is usable ---
    'wsPvt.Unprotect Password:=SYS_SHEET_PWD
    wsPvt.Cells.UnMerge
    
    ' --- remove ghost pivot name if any ---
    DeletePivotAnywhere "ptKPI"
    
    ' IMPORTANT: if Excel still thinks the name exists, generate a unique one
    Dim ptName As String
    ptName = "ptKPI"
    If PivotTableNameExists(ptName) Then
        ptName = "ptKPI_" & Format$(Now, "hhnnss")
    End If
    
    
    Dim ptKPI As PivotTable
    Set ptKPI = pc.CreatePivotTable( _
                TableDestination:=wsPvt.Range("A1"), _
                TableName:="ptKPI")
    
    If ptKPI Is Nothing Then
        Err.Raise vbObjectError + 884, , "ptKPI creation returned Nothing (Wizard)."
    End If
    
    ' If the rest of your code expects the name "ptKPI", rename it when safe
    On Error Resume Next
    ptKPI.name = "ptKPI"
    On Error GoTo EH
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Assert ptKPI fields"
    AssertPivotTableFields ptKPI, Array( _
        "Country", "Region", "Center", "Date", "Month", "MonthName", _
        "Budget", "Actual", "Variance", "AchievementPct" _
    )
    
    With ptKPI
        .ManualUpdate = True
        .ClearAllFilters

        .PivotFields("Country").Orientation = xlPageField
        .PivotFields("Month").Orientation = xlPageField


        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        .AddDataField .PivotFields("Variance"), "Total Variance", xlSum
        

        
        .RowAxisLayout xlTabularRow

        .ManualUpdate = False
    End With
    gStep = "EnsurePivotsSlicersChartsAndTables: Refresh ptKPI"
    ptKPI.RefreshTable
    ForceMonthDisplay ptKPI
    DoEvents

    '==================================================
    ' Calc KPI outputs (very hidden)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Write Calc KPIs"
    wsCalc.Cells.Clear
    wsCalc.Visible = xlSheetVeryHidden
    wsCalc.Range("A1").Value = "KPI outputs (do not edit)"
    wsCalc.Range("A2").Value = "Total Budget":   wsCalc.Range("B2").Formula = "=GETPIVOTDATA(""Total Budget"",Pivots!$A$1)"
    wsCalc.Range("A3").Value = "Total Actual":   wsCalc.Range("B3").Formula = "=GETPIVOTDATA(""Total Actual"",Pivots!$A$1)"
    wsCalc.Range("A4").Value = "Total Variance": wsCalc.Range("B4").Formula = "=GETPIVOTDATA(""Total Variance"",Pivots!$A$1)"
    wsCalc.Range("A5").Value = "Achievement %":  wsCalc.Range("B5").Formula = "=GETPIVOTDATA(""Achievement %"",Pivots!$A$1)"
    
    wsCalc.Range("B2:B4").NumberFormat = "#,##0.00"
    wsCalc.Range("B5").NumberFormat = "0.00%"

    '==================================================
    ' Pivot 2: Country Overall (Pivots!A40)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptCountry"
    Dim ptCountry As PivotTable
    Set ptCountry = pc.CreatePivotTable(TableDestination:=wsPvt.Range("A40"), TableName:="ptCountry")
    With ptCountry
        .ManualUpdate = True
        .ClearAllFilters

        .PivotFields("Country").Orientation = xlRowField
        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .DataFields("Total Budget").NumberFormat = "#,##0"
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        .DataFields("Total Actual").NumberFormat = "#,##0"
        

        
        .RowAxisLayout xlTabularRow

        .ManualUpdate = False
    End With
    gStep = "EnsurePivotsSlicersChartsAndTables: Refresh ptCountry"
    Set ptCountry = GetPivotOrFail(wsPvt, "ptCountry", "Refresh ptCountry")
    ptCountry.RefreshTable
    EnsureAchievementDataField ptCountry
    DoEvents

    'TryFixCountryAchievementAxis wsDash, "chCountry"

    '==================================================
    ' Pivot 3: Month Trend (Pivots!N1)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptMonth"
    Dim ptMonth As PivotTable
    Set ptMonth = pc.CreatePivotTable(TableDestination:=wsPvt.Range("N1"), TableName:="ptMonth")
    
    With ptMonth
        .ManualUpdate = True
        .ClearAllFilters
    
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Country").Orientation = xlPageField
            
        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .DataFields("Total Budget").NumberFormat = "#,##0"

        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        .DataFields("Total Actual").NumberFormat = "#,##0"


        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Refresh ptMonth"
    ptMonth.RefreshTable
    EnsureAchievementDataField ptMonth
    ForceMonthDisplay ptMonth
    DoEvents
    
    ' Sort by Month (date)
    On Error Resume Next
    ptMonth.PivotFields("Month").AutoSort xlAscending, "Month"
    On Error GoTo EH
    
    ' Format the label cells (safe after refresh)
    On Error Resume Next
    If Not ptMonth.PivotFields("Month").LabelRange Is Nothing Then
        ptMonth.PivotFields("Month").LabelRange.NumberFormat = "mmm-yyyy"
    End If
    On Error GoTo EH
    
    ' Clear filters safely (use Month, not MonthName)
    On Error Resume Next
    ptMonth.PivotFields("Month").ClearAllFilters
    ptMonth.PivotFields("Country").ClearAllFilters
    On Error GoTo EH
    
    ptMonth.RefreshTable
    DoEvents
    
    '==================================================
    ' Pivot 4: Center Performance (Pivots!J40)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptCenter"
    Dim ptCenter As PivotTable
    Set ptCenter = pc.CreatePivotTable(TableDestination:=wsPvt.Range("N40"), TableName:="ptCenter")
    With ptCenter
        .ManualUpdate = True
        .ClearAllFilters

        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Center").Orientation = xlRowField
        .PivotFields("Country").Orientation = xlPageField
        .PivotFields("Month").Orientation = xlPageField
        

        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False
    End With
    gStep = "EnsurePivotsSlicersChartsAndTables: Refresh ptCenter"
    ptCenter.RefreshTable
    ForceMonthDisplay ptCenter

    '==================================================
    ' Pivot 5: Country x Month (Pivots!A90)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptCM"
    Dim ptCM As PivotTable
    Dim cmRow As Long
    cmRow = NextFreePivotRow(wsPvt, 90) ' at least 90, but move down if needed
    
    
    Set ptCM = pc.CreatePivotTable(TableDestination:=wsPvt.Cells(cmRow, 1), TableName:="ptCountryMonth")

    
    With ptCM
        .ManualUpdate = True
        .ClearAllFilters
        .PivotFields("Country").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With
    
    ptCM.RefreshTable
    ForceMonthDisplay ptCM
    DoEvents
    
    ' Next pivot starts AFTER ptCM footprint
    Dim nextStartRow As Long
    nextStartRow = ptCM.TableRange2.Row + ptCM.TableRange2.Rows.Count + 3
    
    '==================================================
    ' Pivot 6: Country x Center (Pivots!A???)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Build ptCC"
    Dim ptCC As PivotTable
    Set ptCC = pc.CreatePivotTable(TableDestination:=wsPvt.Cells(nextStartRow, 1), TableName:="ptCountryCenter")
    
    With ptCC
        .ManualUpdate = True
        .ClearAllFilters
        .PivotFields("Country").Orientation = xlPageField
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Center").Orientation = xlRowField
        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum
        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With
    
    ptCC.RefreshTable
    DoEvents

    gStep = "EnsurePivotsSlicersChartsAndTables: Clear ptMonth filters (post pivots)"
    On Error Resume Next
    ptMonth.PivotFields("Month").ClearAllFilters
    ptMonth.PivotFields("Country").ClearAllFilters
    On Error GoTo EH
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Refresh ptMonth (post pivots)"
    ptMonth.RefreshTable
    DoEvents

    '==================================================
    ' Slicers (Dashboard) and connections
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: Create slicer caches"
    Dim scCountry As SlicerCache, scMonth As SlicerCache
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Assert all pivots share cache"
    AssertSamePivotCache ptKPI, ptCountry, gStep
    AssertSamePivotCache ptKPI, ptMonth, gStep
    AssertSamePivotCache ptKPI, ptCenter, gStep
    AssertSamePivotCache ptKPI, ptCM, gStep
    AssertSamePivotCache ptKPI, ptCC, gStep
    
    Set scCountry = ThisWorkbook.SlicerCaches.Add2(ptKPI, "Country", "scCountry")
    
    gStep = "EnsurePivotsSlicersChartsAndTables: Assert all pivots share cache"
    AssertSamePivotCache ptKPI, ptCountry, gStep
    AssertSamePivotCache ptKPI, ptMonth, gStep
    AssertSamePivotCache ptKPI, ptCenter, gStep
    AssertSamePivotCache ptKPI, ptCM, gStep
    AssertSamePivotCache ptKPI, ptCC, gStep
    Set scMonth = ThisWorkbook.SlicerCaches.Add2(ptKPI, "Month", "scMonth")

    gStep = "EnsurePivotsSlicersChartsAndTables: Connect slicers to pivots"
    ' Country slicer controls all pivots
    scCountry.PivotTables.AddPivotTable ptKPI
    scCountry.PivotTables.AddPivotTable ptCountry
    scCountry.PivotTables.AddPivotTable ptMonth
    scCountry.PivotTables.AddPivotTable ptCenter
    scCountry.PivotTables.AddPivotTable ptCM
    scCountry.PivotTables.AddPivotTable ptCC

    ' Month slicer controls month-aware pivots
    scMonth.PivotTables.AddPivotTable ptKPI
    scMonth.PivotTables.AddPivotTable ptCountry
    scMonth.PivotTables.AddPivotTable ptMonth
    scMonth.PivotTables.AddPivotTable ptCenter
    scMonth.PivotTables.AddPivotTable ptCM
    scMonth.PivotTables.AddPivotTable ptCC

    gStep = "EnsurePivotsSlicersChartsAndTables: Add slicers to Dashboard"
    Dim slCountry As Slicer, slMonth As Slicer
    Set slCountry = scCountry.Slicers.Add(wsDash, , "slCountry", "Country", 20, 60, 220, 140)
    Set slMonth = scMonth.Slicers.Add(wsDash, , "slMonth", "Month", 20, 320, 260, 140)
    On Error Resume Next
    slCountry.Shape.name = "slCountry"
    slMonth.Shape.name = "slMonth"
    On Error GoTo 0

    '==================================================
    ' PivotCharts (Dashboard) - safe binding (no SetSourceData)
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: PivotCharts - chMonth"
    Dim shMonth As Shape, shCountry As Shape, shCenter As Shape
    ' --- ensure ptMonth has visible items before chart binding ---
    On Error Resume Next
    ptMonth.ManualUpdate = False
    ptMonth.PivotFields("Month").ClearAllFilters
    ptMonth.PivotFields("Country").ClearAllFilters
    On Error GoTo EH
    
    ptMonth.RefreshTable
    DoEvents

    ptCountry.ManualUpdate = False
    ptCountry.RefreshTable
    DoEvents

    ' Build/rebuild main dashboard charts as true PivotCharts.
    EnsureMainDashboardPivotCharts wsDash, wsPvt, ptMonth, ptCountry, True

    'gStep = "EnsurePivotsSlicersChartsAndTables: PivotCharts - chCenter"
    'Set shCenter = AddPivotChartClean(wsDash, wsPvt, ptCenter, xlColumnClustered, "Center Performance (Selected Country)", 1120, 200, 520, 260, "chCenter")

    '==================================================
    ' Breakdown tables (Dashboard) via GETPIVOTDATA
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: RenderCountryMonthTable"
'    RenderCountryMonthTable wsDash, ptCM, ptCM.TableRange2.Cells(1, 1)
    RenderCountryMonthTable wsDash, ptCM, ptCM.TableRange1.Cells(1, 1)

    gStep = "EnsurePivotsSlicersChartsAndTables: RenderCountryCenterTable"
'    RenderCountryCenterTable wsDash, ptCC, ptCC.TableRange2.Cells(1, 1)
    RenderCountryCenterTable wsDash, ptCC, ptCC.TableRange1.Cells(1, 1)

    '==================================================
    ' Exec formatting + button + protection
    '==================================================
    gStep = "EnsurePivotsSlicersChartsAndTables: FormatDashboard"
    FormatDashboard wsDash

    gStep = "EnsurePivotsSlicersChartsAndTables: FormatPivotForExec ptKPI"
    FormatPivotForExec ptKPI


    gStep = "EnsurePivotsSlicersChartsAndTables: FormatPivotForExec ptMonth"
    FormatPivotForExec ptMonth

    gStep = "EnsurePivotsSlicersChartsAndTables: FormatPivotForExec ptCenter"
    FormatPivotForExec ptCenter

    gStep = "EnsurePivotsSlicersChartsAndTables: FormatPivotForExec ptCM"
    FormatPivotForExec ptCM

    gStep = "EnsurePivotsSlicersChartsAndTables: FormatPivotForExec ptCC"
    FormatPivotForExec ptCC

    gStep = "EnsurePivotsSlicersChartsAndTables: AddImportButton"
    AddImportButton wsDash

    gStep = "EnsurePivotsSlicersChartsAndTables: ProtectDashboard"
    'ProtectDashboard wsDash

    ' Apply combo formatting to all charts after everything is built
    ' Inject Worksheet_PivotTableUpdate into Pivots sheet module (survives re-import)
    'InjectPivotSheetEvents

    ApplyAllComboFormats

    ' Delete temporary/scratch sheets left by chart or pivot building
    DeleteScratchSheets

    gStep = "EnsurePivotsSlicersChartsAndTables: done"

    Exit Sub
EH:
    Err.Raise vbObjectError + 999, , gStep & ": " & Err.Number & " - " & Err.Description

End Sub

'==================================================
' DASHBOARD BREAKDOWN TABLES (SAFE: use PivotItems)
'==================================================

Private Sub RenderCountryMonthTable(ByVal wsDash As Worksheet, ByVal ptCM As PivotTable, ByVal ptAnchor As Range)
    gStep = "RenderCountryMonthTable: start"

    Dim topRow As Long: topRow = 45
    Dim leftCol As Long: leftCol = 1 'A

    wsDash.Range("A45:Z69").ClearContents

    wsDash.Cells(topRow, leftCol).Value = "Country x Month Breakdown"
    wsDash.Cells(topRow, leftCol).Font.Bold = True

    Dim pfCountry As PivotField, pfMonth As PivotField
    gStep = "RenderCountryMonthTable: get PivotFields"
    Set pfCountry = ptCM.PivotFields("Country")

    Set pfMonth = ptCM.PivotFields("Month")

    gStep = "RenderCountryMonthTable: write month headers"
    '----- Header row (months)
    wsDash.Cells(topRow + 1, leftCol).Value = "Country"
    wsDash.Cells(topRow + 1, leftCol).Font.Bold = True

    Dim m As Long, outCol As Long
    outCol = leftCol + 1

    Dim piM As PivotItem
    For Each piM In pfMonth.PivotItems
        If piM.name <> "(blank)" Then
    
            Dim dt As Date
            dt = PivotItemToDate(piM)
    
            wsDash.Cells(topRow + 1, outCol).Value = dt
            wsDash.Cells(topRow + 1, outCol).NumberFormat = "mmm-yyyy"

            wsDash.Cells(topRow + 1, outCol).Font.Bold = True
    
            outCol = outCol + 1
            If outCol > 26 Then Exit For ' stop at Z
        End If
    Next piM


    Dim lastMonthCol As Long
    lastMonthCol = outCol - 1
    If lastMonthCol < leftCol + 1 Then
        gStep = "RenderCountryMonthTable: no months found, exit"
        Exit Sub ' no months
    End If

    '----- Budget block label
    Dim outRow As Long
    outRow = topRow + 2
    wsDash.Cells(outRow, leftCol).Value = "Budget"
    wsDash.Cells(outRow, leftCol).Font.Bold = True
    outRow = outRow + 1

    gStep = "RenderCountryMonthTable: write Budget block formulas"
    '----- Budget rows by country
    Dim piC As PivotItem
    For Each piC In pfCountry.PivotItems
        If piC.name <> "(blank)" Then
            wsDash.Cells(outRow, leftCol).Value = piC.name

            Dim c As Long
            For c = leftCol + 1 To lastMonthCol
                Dim mName As String
                mName = CStr(wsDash.Cells(topRow + 1, c).Value)
                wsDash.Cells(outRow, c).Formula = _
                    "=IFERROR(GETPIVOTDATA(""Total Budget""," & ptAnchor.Address(True, True, xlA1, True) & _
                    ",""Country"",""" & EscapeQuotes(piC.name) & """,""Month""," & wsDash.Cells(topRow + 1, c).Address(False, False) & "),0)"
            Next c

            outRow = outRow + 1
            If outRow > topRow + 30 Then Exit For
        End If
    Next piC

    outRow = outRow + 1 ' spacer

    '----- Actual block label
    wsDash.Cells(outRow, leftCol).Value = "Actual"
    wsDash.Cells(outRow, leftCol).Font.Bold = True
    outRow = outRow + 1

    gStep = "RenderCountryMonthTable: write Actual block formulas"
    '----- Actual rows by country
    For Each piC In pfCountry.PivotItems
        If piC.name <> "(blank)" Then
            wsDash.Cells(outRow, leftCol).Value = piC.name

            For c = leftCol + 1 To lastMonthCol
                
                wsDash.Cells(outRow, c).Formula = _
                    "=IFERROR(GETPIVOTDATA(""Total Actual""," & ptAnchor.Address(True, True, xlA1, True) & _
                    ",""Country"",""" & EscapeQuotes(piC.name) & """,""Month""," & wsDash.Cells(topRow + 1, c).Address(False, False) & "),0)"
            
            Next c

            outRow = outRow + 1
            If outRow > topRow + 60 Then Exit For
        End If
    Next piC

    gStep = "RenderCountryMonthTable: apply formats"
    ' formats
    wsDash.Range(wsDash.Cells(topRow + 2, leftCol + 1), wsDash.Cells(topRow + 65, lastMonthCol)).NumberFormat = "#,##0"

    ' Apply borders (header + blocks)
    Dim lastRow As Long
    lastRow = outRow - 1
    
    ApplyTableBorders wsDash.Range(wsDash.Cells(topRow, leftCol), wsDash.Cells(lastRow, lastMonthCol))

    gStep = "RenderCountryMonthTable: done"
End Sub
Private Sub RenderCountryCenterTable(ByVal wsDash As Worksheet, ByVal ptCC As PivotTable, ByVal ptAnchor As Range)
    gStep = "RenderCountryCenterTable: start"

    Dim topRow As Long: topRow = 70
    Dim leftCol As Long: leftCol = 1 'A

    wsDash.Range("A70:H120").ClearContents

    wsDash.Cells(topRow, leftCol).Value = "Country x Center Breakdown"
    wsDash.Cells(topRow, leftCol).Font.Bold = True

    ' Headers
    wsDash.Cells(topRow + 1, leftCol).Value = "Region"
    wsDash.Cells(topRow + 1, leftCol + 1).Value = "Center"
    wsDash.Cells(topRow + 1, leftCol + 2).Value = "Total Budget"
    wsDash.Cells(topRow + 1, leftCol + 3).Value = "Total Actual"
    wsDash.Cells(topRow + 1, leftCol + 4).Value = "Ach %"
    wsDash.Range(wsDash.Cells(topRow + 1, leftCol), wsDash.Cells(topRow + 1, leftCol + 4)).Font.Bold = True

    ' Ensure pivot is refreshed
    ptCC.ManualUpdate = False
    ptCC.RefreshTable
    DoEvents

    Dim rng As Range
    Set rng = ptCC.TableRange1

    ' Pivot label columns in the output (usually first two)
    Dim regionCol As Long: regionCol = 1
    Dim centerCol As Long: centerCol = 2

    '==================================================
    ' Find the TRUE header row inside the pivot output:
    ' a row that contains "Region" and "Center"
    '==================================================
    Dim headerRow As Long: headerRow = 0
    Dim r As Long

    For r = 1 To rng.Rows.Count
        Dim h1 As String, h2 As String
        h1 = LCase$(Trim$(CStr(rng.Cells(r, regionCol).Value)))
        h2 = LCase$(Trim$(CStr(rng.Cells(r, centerCol).Value)))

        If h1 = "region" And h2 = "center" Then
            headerRow = r
            Exit For
        End If
    Next r

    If headerRow = 0 Then
        ' Fallback: if pivot uses different captions, just avoid report filter rows
        headerRow = 1
    End If

    Dim outRow As Long: outRow = topRow + 2
    Dim lastRegion As String: lastRegion = ""

    ' Start reading AFTER the header row -> avoids Country/Singapore filter rows
    For r = headerRow + 1 To rng.Rows.Count
        Dim vRegion As String, vCenter As String
        vRegion = Trim$(CStr(rng.Cells(r, regionCol).Value))
        vCenter = Trim$(CStr(rng.Cells(r, centerCol).Value))

        ' Stop on grand total
        If LCase$(vRegion) = "grand total" Or LCase$(vCenter) = "grand total" Then Exit For

        ' Skip empty
        If vRegion = "" And vCenter = "" Then GoTo NextR

        ' Region group row (Center blank)
        If vRegion <> "" And vCenter = "" Then
            wsDash.Cells(outRow, leftCol).Value = vRegion
            wsDash.Cells(outRow, leftCol).Font.Bold = True
            lastRegion = vRegion
            outRow = outRow + 1
            GoTo NextR
        End If

        ' Carry down last region if repeating labels are off
        If vRegion <> "" Then lastRegion = vRegion
        If vRegion = "" Then vRegion = lastRegion

        ' Still skip any accidental header remnants
        If LCase$(vRegion) = "region" Or LCase$(vCenter) = "center" Then GoTo NextR
        If LCase$(vRegion) = "country" Or LCase$(vCenter) = "country" Then GoTo NextR

        ' Write row
        wsDash.Cells(outRow, leftCol).Value = vRegion
        wsDash.Cells(outRow, leftCol + 1).Value = vCenter

        wsDash.Cells(outRow, leftCol + 2).Formula = _
            "=IFERROR(GETPIVOTDATA(""Total Budget""," & ptAnchor.Address(True, True, xlA1, True) & _
            ",""Region"",""" & EscapeQuotes(vRegion) & """,""Center"",""" & EscapeQuotes(vCenter) & """),0)"

        wsDash.Cells(outRow, leftCol + 3).Formula = _
            "=IFERROR(GETPIVOTDATA(""Total Actual""," & ptAnchor.Address(True, True, xlA1, True) & _
            ",""Region"",""" & EscapeQuotes(vRegion) & """,""Center"",""" & EscapeQuotes(vCenter) & """),0)"

        wsDash.Cells(outRow, leftCol + 4).Formula = _
            "=IFERROR(GETPIVOTDATA(""Ach %""," & ptAnchor.Address(True, True, xlA1, True) & _
            ",""Region"",""" & EscapeQuotes(vRegion) & """,""Center"",""" & EscapeQuotes(vCenter) & """),0)"

        outRow = outRow + 1
        If outRow > topRow + 60 Then Exit For

NextR:
    Next r

    wsDash.Range(wsDash.Cells(topRow + 2, leftCol + 2), wsDash.Cells(outRow - 1, leftCol + 3)).NumberFormat = "#,##0"
    wsDash.Range(wsDash.Cells(topRow + 2, leftCol + 4), wsDash.Cells(outRow - 1, leftCol + 4)).NumberFormat = "0.00%"

    ApplyTableBorders wsDash.Range(wsDash.Cells(topRow, leftCol), wsDash.Cells(outRow - 1, leftCol + 4))

    gStep = "RenderCountryCenterTable: done"
End Sub





Private Function EscapeQuotes(ByVal s As String) As String
    gStep = "EscapeQuotes"
    EscapeQuotes = Replace(s, """", """""")
End Function


'=========================
' CHART / BUTTON / CLEANUP HELPERS
'=========================

'----------------------------------------------------------------
' Applies clustered-column + line combo format to a PivotChart.
' Series matching "ctrach" -> line on secondary axis (CtrAch field).
' All others -> clustered column on primary axis.
' Called once at build time. PivotChart preserves series types
' across refreshes when data field order stays constant.
'----------------------------------------------------------------
Private Sub ApplyPivotChartComboFormat(ByVal ch As Chart)
    On Error Resume Next

    ch.chartType = xlColumnClustered
    DoEvents

    Dim s As Series
    Dim hasSecondary As Boolean: hasSecondary = False
    For Each s In ch.SeriesCollection
        ' Primary detection by series name/axis.
        If IsAchievementSeries(s) Then
            s.AxisGroup = xlSecondary
            hasSecondary = True
        Else
            s.AxisGroup = xlPrimary
        End If
        Err.Clear
    Next s

    ' Fallback: if Ach series name was not recognized, assume 3rd series is Ach%.
    If (Not hasSecondary) And ch.SeriesCollection.Count >= 3 Then
        ch.SeriesCollection(3).AxisGroup = xlSecondary
        hasSecondary = True
    End If

    If hasSecondary Then
        ch.HasAxis(xlValue, xlSecondary) = True
        Err.Clear
        DoEvents
    End If

    For Each s In ch.SeriesCollection
        Dim isAch As Boolean
        isAch = IsAchievementSeries(s)
        If (Not isAch) And hasSecondary And ch.SeriesCollection.Count >= 3 Then
            If s.PlotOrder = 3 Then isAch = True
        End If

        If isAch Then
            s.chartType = xlLineMarkers
            s.AxisGroup = xlSecondary
            s.HasDataLabels = True
            s.DataLabels.ShowValue = True
            s.DataLabels.NumberFormat = "0.0%"
            s.DataLabels.Position = xlLabelPositionAbove
            s.DataLabels.Font.size = 9
        Else
            s.chartType = xlColumnClustered
            s.AxisGroup = xlPrimary
            s.HasDataLabels = True
            s.DataLabels.ShowValue = True
            s.DataLabels.NumberFormat = "#,##0"
            s.DataLabels.Font.size = 8
        End If
        Err.Clear
    Next s

    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    If hasSecondary Then
        With ch.Axes(xlValue, xlSecondary)
            .TickLabels.NumberFormat = "0%"
            .TickLabelPosition = xlNextToAxis
            .Format.Line.Visible = msoTrue
            .HasTitle = False
        End With
    End If
    On Error Resume Next
    ch.HasAxis(xlCategory, xlSecondary) = False
    ch.ChartGroups(1).GapWidth = 80
    AutoScalePercentAxis ch
    Err.Clear
End Sub

Private Function AddPivotChartClean( _
    ByVal wsDash As Worksheet, _
    ByVal wsPvt As Worksheet, _
    ByVal pt As PivotTable, _
    ByVal chartType As XlChartType, _
    ByVal title As String, _
    ByVal left As Double, _
    ByVal top As Double, _
    ByVal width As Double, _
    ByVal height As Double, _
    ByVal chartName As String _
) As Shape
    ' PivotChart — Budget + Actual columns, CtrAch line on secondary axis.
    ' pt has 3 data fields: Total Budget, Total Actual, CtrAch.
    ' Slicers update pivot -> chart redraws automatically. Zero VBA on slicer click.

    gStep = "AddPivotChartClean: start [" & chartName & "]"
    On Error GoTo CHART_FAIL

    On Error Resume Next
    wsDash.Shapes(chartName).Delete
    On Error GoTo CHART_FAIL

    pt.ManualUpdate = False
    pt.RefreshTable
    DoEvents

    If pt.DataBodyRange Is Nothing Then
        Dim shpEmpty As Shape
        Set shpEmpty = wsDash.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
        shpEmpty.name = chartName
        shpEmpty.TextFrame2.TextRange.Text = title & vbCrLf & "(No data)"
        Set AddPivotChartClean = shpEmpty
        Exit Function
    End If

    ' Create PivotChart
    Dim shp As Shape
    Set shp = wsDash.Shapes.AddChart(xlColumnClustered, left, top, width, height)
    If shp Is Nothing Then GoTo CHART_FAIL
    shp.name = chartName

    Dim ch As Chart: Set ch = shp.Chart

    ' Bind to pivot
    On Error Resume Next
    ch.PivotLayout.PivotTable = pt
    If Err.Number <> 0 Then
        Err.Clear
        ch.SetSourceData Source:=pt.TableRange2
    End If
    On Error GoTo CHART_FAIL

    ch.HasTitle = True
    ch.chartTitle.Text = title
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight
    ch.ShowAllFieldButtons = False
    ch.ShowLegendFieldButtons = False
    ch.ShowAxisFieldButtons = False
    ch.ShowValueFieldButtons = False
    ch.ShowReportFilterFieldButtons = False

    ' Apply combo format: Budget+Actual = clustered columns primary, CtrAch = line secondary
    ApplyPivotChartComboFormat ch

    ' Log actual series names for diagnosis
    On Error Resume Next
    Dim si As Long
    For si = 1 To ch.SeriesCollection.Count
        LogMsg "[APC] series(" & si & ")=[" & ch.SeriesCollection(si).name & "] axisGroup=" & ch.SeriesCollection(si).AxisGroup
    Next si
    Err.Clear

    LogMsg "[APC] built PivotChart=" & chartName & " series=" & ch.SeriesCollection.Count
    Set AddPivotChartClean = shp
    Exit Function

CHART_FAIL:
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    Dim shpFail As Shape
    Set shpFail = wsDash.Shapes.AddShape(msoShapeRectangle, left, top, width, height)
    shpFail.name = chartName
    shpFail.TextFrame2.TextRange.Text = title & vbCrLf & "(Chart build failed: " & Err.Number & " " & Err.Description & ")"
    Set AddPivotChartClean = shpFail
    LogMsg "[APC] CHART_FAIL err=" & Err.Number & " " & Err.Description
End Function

'----------------------------------------------------------------
' Applies clustered-column + line combo format to a PivotChart.
' Series matching "ach" or "%" -> line on secondary axis.
' All others -> clustered column on primary axis.
' Safe to call at build time. Does NOT need to be called on slicer
' update because PivotChart preserves series types across refreshes
' when the data field order in the pivot stays constant.
'----------------------------------------------------------------

'----------------------------------------------------------------
' Writes chart helper range to wsDash cols AZ:BC (cols 52-55).
' Layout:  AZ=Center  BA=Total Budget  BB=Total Actual  BC=Ach%
'   Row 1:  headers
'   Row 2+: one row per visible center in current pivot state
' Called at build time and from RefreshCountryDashboardChart (OnTime).
' The plain static chart reads directly from these cells via SERIES
' formulas - so slicer changes just update cell values, zero chart
' object writes needed in OnTime (immune to all pivot cache locks).
'----------------------------------------------------------------

Private Sub AddImportButton(ws As Worksheet)
    gStep = "AddImportButton: start"

    Dim shp As Shape
    Dim shpPdf As Shape

    ' Delete ALL duplicates (both current + legacy names)
    Call DeleteAllShapesByName(ws, "btnImportRefresh")
    Call DeleteAllShapesByName(ws, "btnRegeneratePdf")
    Call DeleteAllShapesByName(ws, "btnPrevMthAdj")
    Call DeleteAllShapesByName(ws, "btnEnableDeleteOptionForWorksheet") ' legacy bad name

    ' Import & Refresh
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 620, 30, 180, 60)
    shp.name = "btnImportRefresh"
    shp.TextFrame2.TextRange.Text = "Import & Refresh"
    shp.TextFrame2.TextRange.Font.size = 14
    shp.Fill.ForeColor.RGB = RGB(37, 99, 235)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "ImportAndRefresh"
    shp.ZOrder msoBringToFront

    ' Regenerate PDF
    Set shpPdf = ws.Shapes.AddShape(msoShapeRoundedRectangle, 850, 30, 210, 60)
    shpPdf.name = "btnRegeneratePdf"
    shpPdf.TextFrame2.TextRange.Text = "Regenerate Pdf"
    shpPdf.TextFrame2.TextRange.Font.size = 14
    shpPdf.Fill.ForeColor.RGB = RGB(37, 99, 235)
    shpPdf.Line.Visible = msoFalse
    shpPdf.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpPdf.OnAction = "ExportAllCharts_Fast_ChartSheets_Direct"
    shpPdf.ZOrder msoBringToFront

    
    ' Monthly Adjustment
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 620, 100, 180, 60)
    shp.name = "btnPrevMthAdj"
    shp.TextFrame2.TextRange.Text = "Prev Mth Adj"
    shp.TextFrame2.TextRange.Font.size = 14
    shp.Fill.ForeColor.RGB = RGB(37, 99, 235)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "ImportPrevMonthAdj"
    shp.ZOrder msoBringToFront
    
    gStep = "AddImportButton: done"
End Sub
Private Sub ClearSheet(ByVal ws As Worksheet)
    gStep = "ClearSheet: start [" & ws.name & "]"

    On Error Resume Next

    ' 1) Delete pivot objects (clearing range is NOT enough)
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    For Each pt In ws.PivotTables
        pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        pt.TableRange2.Clear
    Next pt

    ' Delete ALL pivot tables by recreating sheet content
    ' (Excel does not expose a PivotTable.Delete, but removing the sheet content + clearing pivot caches works reliably)
    ws.Cells.Clear
    ws.Cells.ClearFormats

    On Error GoTo 0
    gStep = "ClearSheet: done [" & ws.name & "]"
End Sub

Private Sub ClearDashboardObjects(ByVal ws As Worksheet)
    gStep = "ClearDashboardObjects: start [" & ws.name & "]"

    On Error Resume Next

    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Dim nm As String
        nm = ws.Shapes(i).name

        ' Delete our generated items only
        If ws.Shapes(i).Type = msoSlicer Then
            ws.Shapes(i).Delete
        ElseIf left$(nm, 2) = "ch" Then
            ws.Shapes(i).Delete
        ElseIf nm = "btnImportRefresh" Then
            ws.Shapes(i).Delete
        ElseIf nm = "btnPrevMthAdj" Then
            ws.Shapes(i).Delete
        ElseIf nm = "btnRegeneratePdf" Or nm = "btnEnableDeleteOptionForWorksheet" Then
            ws.Shapes(i).Delete
        End If
    Next i

    ' Also remove legacy chartobjects if any
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    On Error GoTo 0

    gStep = "ClearDashboardObjects: done [" & ws.name & "]"
End Sub




'Private Sub DeleteSlicerCacheIfExists(ByVal cacheName As String)
'    gStep = "DeleteSlicerCacheIfExists: " & cacheName'

'    On Error Resume Next
'    ThisWorkbook.SlicerCaches(cacheName).Delete
'    On Error GoTo 0
'End Sub


'=========================
' EXEC FORMATTING + PROTECTION
'=========================
Private Sub FormatDashboard(ByVal ws As Worksheet)
    gStep = "FormatDashboard: start"

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.size = 11
    On Error GoTo 0

    gStep = "FormatDashboard: done"
End Sub


Private Sub FormatPivotForExec(ByVal pt As PivotTable)
    gStep = "FormatPivotForExec: start [" & pt.name & "]"

    On Error Resume Next

    pt.ManualUpdate = True
    pt.RowAxisLayout xlTabularRow
    pt.ColumnGrand = False
    pt.RowGrand = False

    Dim pf As PivotField
    For Each pf In pt.RowFields
        Dim i As Integer
        For i = 1 To 12
            pf.Subtotals(i) = False
        Next i
    Next pf

    pt.TableStyle2 = "PivotStyleLight16"
    pt.ManualUpdate = False
    pt.RefreshTable

    On Error GoTo 0

    gStep = "FormatPivotForExec: done [" & pt.name & "]"
End Sub


Private Sub ProtectDashboard(ByVal ws As Worksheet)
    gStep = "ProtectDashboard: start"

    On Error Resume Next
    'ws.Unprotect Password:=SYS_SHEET_PWD

    On Error GoTo 0

    gStep = "ProtectDashboard: done"
End Sub

Private Sub ClearPivotTablesOnSheet(ByVal ws As Worksheet)
    gStep = "ClearPivotTablesOnSheet: start [" & ws.name & "]"

    On Error Resume Next

    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt

    On Error GoTo 0

    gStep = "ClearPivotTablesOnSheet: done [" & ws.name & "]"
End Sub

Private Sub DeleteAllPivotTablesOnSheet(ByVal ws As Worksheet)
    gStep = "DeleteAllPivotTablesOnSheet: start [" & ws.name & "]"

    On Error Resume Next

    Dim pt As PivotTable
    Dim rng As Range

    ' Collect and clear each pivot footprint
    For Each pt In ws.PivotTables
        Set rng = pt.TableRange2
        rng.Clear
    Next pt

    On Error GoTo 0

    gStep = "DeleteAllPivotTablesOnSheet: done [" & ws.name & "]"
End Sub


Public Function BuildKey(ByVal country As String, ByVal region As String, ByVal center As String, ByVal d As Date) As String
    BuildKey = LCase$(Trim$(country)) & "|" & LCase$(Trim$(region)) & "|" & LCase$(Trim$(center)) & "|" & Format$(d, "yyyymmdd")
End Function

Public Function BuildExistingKeyIndex(ByVal lo As ListObject) As Object
    gStep = "BuildExistingKeyIndex: start"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If lo.DataBodyRange Is Nothing Then
        Set BuildExistingKeyIndex = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim rw As Range
        Set rw = lo.ListRows(i).Range

        Dim vCountry As String, vRegion As String, vCenter As String
        Dim vDate As Variant

        vCountry = Trim$(CStr(rw.Cells(1, 1).Value))  ' Country
        vRegion = Trim$(CStr(rw.Cells(1, 2).Value))   ' Region
        vCenter = Trim$(CStr(rw.Cells(1, 3).Value))   ' Center
        vDate = rw.Cells(1, 4).Value                  ' Date

        If Len(vCountry) > 0 And Len(vRegion) > 0 And Len(vCenter) > 0 And IsDate(vDate) Then
            Dim k As String
            k = BuildKey(vCountry, vRegion, vCenter, CDate(vDate))
            dict(k) = i   ' ListRow index
        End If
    Next i

    gStep = "BuildExistingKeyIndex: done"
    Set BuildExistingKeyIndex = dict
End Function

Public Sub UpsertRow(ByVal lo As ListObject, ByVal keyIndex As Object, _
                      ByVal country As String, ByVal region As String, ByVal center As String, _
                      ByVal d As Date, ByVal m As Date, ByVal mName As String, _
                      ByVal budget As Double, ByVal actual As Double, ByVal variance As Double, _
                      ByVal ach As Variant)

    gStep = "UpsertRow: start [" & country & "|" & center & "|" & Format$(d, "yyyy-mm-dd") & "]"

    Dim k As String
    k = BuildKey(country, region, center, d)

    Dim rowRange As Range
    Dim lr As ListRow
    
    If keyIndex.Exists(k) Then
        Set lr = lo.ListRows(CLng(keyIndex(k)))
    Else
        Set lr = lo.ListRows.Add
        keyIndex(k) = lr.Index
    End If
    
    Set rowRange = lr.Range

    ' Write values
    With rowRange
        .Cells(1, 1).Value = UCase$(country)
        .Cells(1, 2).Value = region
        .Cells(1, 3).Value = center
        .Cells(1, 4).Value = d
        .Cells(1, 5).Value = m
        .Cells(1, 6).Value = mName
        .Cells(1, 7).Value = budget
        .Cells(1, 8).Value = actual
        .Cells(1, 9).Value = variance
        If IsEmpty(ach) Then
            If Not IsEmpty(.Cells(1, 10).Value) Then
                .Cells(1, 10).ClearContents
            End If
        Else
            .Cells(1, 10).Value = ach
        End If
        ' Upload metadata for this import run
        Dim iUploadId As Long, iUploadedAt As Long
        iUploadId = lo.ListColumns("UploadId").Index
        iUploadedAt = lo.ListColumns("UploadedAt").Index
        
        .Cells(1, iUploadId).Value = gUploadId
        .Cells(1, iUploadedAt).Value = gUploadedAt
'        .Cells(1, 11).Value = gUploadId
'        .Cells(1, 12).Value = gUploadedAt
    End With

    gStep = "UpsertRow: done"
End Sub


Private Sub RemoveBlankRowsInTblData()
    gStep = "RemoveBlankRowsInTblData: start"

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("Data").ListObjects("tblData")

    If lo.DataBodyRange Is Nothing Then
        gStep = "RemoveBlankRowsInTblData: no DataBodyRange, exit"
        Exit Sub
    End If

    Dim r As Long
    For r = lo.DataBodyRange.Rows.Count To 1 Step -1
        Dim c As String, ctr As String, v As Variant
        c = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value))
        ctr = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value))
        v = lo.DataBodyRange.Cells(r, 4).Value

        If Len(c) = 0 Or Len(ctr) = 0 Or Not IsDate(v) Then
            gStep = "RemoveBlankRowsInTblData: delete row " & r
            lo.ListRows(r).Delete
        End If
    Next r

    gStep = "RemoveBlankRowsInTblData: done"
End Sub

'==================================================
' One Dashboard + one Pivots sheet per country
'==================================================

Private Sub BuildCountryDashboards()
    gStep = "BuildCountryDashboards: start"

    BuildRunBegin

    On Error GoTo EH

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo CleanExit

    ' Validate schema headers exist (you already have this)
    AssertTblDataHeaders lo

    Dim countries As Collection
    Set countries = GetUniqueCountries(lo)

    Dim country As Variant
    For Each country In countries
        Dim c As String
        c = Trim$(CStr(country))

        If Len(c) > 0 Then
            If BuildCountryOncePerRun(c) Then
                Debug.Print "BuildOneCountryDashboard: " & c & " RunId=" & gBuildRunId
                BuildOneCountryDashboard c, lo   ' <-- IMPORTANT: pass ListObject, not shared PivotCache
            End If
        End If
    Next country

CleanExit:
    BuildRunEnd
    gStep = "BuildCountryDashboards: done"
    Exit Sub

EH:
    BuildRunEnd
    Err.Raise vbObjectError + 7702, , gStep & ": " & Err.Number & " - " & Err.Description
End Sub


'==================================================
' Country Dashboard: static combo chart (Budget/Actual columns + Ach% line)
' Same architecture as main dashboard chMonth/chCountry.
'==================================================

Public Function RefreshCountryChartData(ByVal wsPvt As Worksheet, ByVal wsDash As Worksheet, ByVal pt As PivotTable, _
                                    Optional ByVal forceRefresh As Boolean = False, _
                                    Optional ByVal startCell As String = "AA1") As Long
    ' Returns iOut = number of data rows written (excluding header).
    ' Callers must use .Resize(iOut+1, 5) instead of .CurrentRegion to avoid
    ' bleeding into adjacent 200-row blocks.
    On Error GoTo SafeExit
    ' Write to Dashboard sheet at startCell - each region gets its own block
    Dim pasteCell As Range
    Set pasteCell = wsDash.Range(startCell)
    pasteCell.Resize(200, 5).Clear
    LogMsg "[RCD2] start, pt=" & pt.name & " forceRefresh=" & forceRefresh

    If forceRefresh Then
        On Error Resume Next
        pt.ManualUpdate = False
        pt.RefreshTable
        On Error GoTo SafeExit
    End If

    If pt.DataBodyRange Is Nothing Then
        LogMsg "[RCD2] DataBodyRange is Nothing, exit"
        GoTo SafeExit
    End If
    LogMsg "[RCD2] DataBodyRange rows=" & pt.DataBodyRange.Rows.Count & " firstRow=" & pt.DataBodyRange.Row

    ' Log all pivot fields to understand structure
    Dim pf As PivotField
    On Error Resume Next
    Dim pfList As String: pfList = ""
    For Each pf In pt.PivotFields
        pfList = pfList & pf.name & "(" & pf.Orientation & ") "
    Next pf
    LogMsg "[RCD2] fields: " & left(pfList, 200)
    On Error GoTo SafeExit

    ' Get pivot fields
    Dim rfRegion As PivotField, rfCenter As PivotField
    Dim dfBgt As PivotField, dfAct As PivotField
    On Error Resume Next
    Set rfRegion = pt.PivotFields("Region")
    Set rfCenter = pt.PivotFields("Center")
    Set dfBgt = pt.PivotFields("Total Budget")
    Set dfAct = pt.PivotFields("Total Actual")
    On Error GoTo SafeExit
    LogMsg "[RCD2] rfRegion=" & (Not rfRegion Is Nothing) & " rfCenter=" & (Not rfCenter Is Nothing) & " dfBgt=" & (Not dfBgt Is Nothing)

    ' Log visible region items
    On Error Resume Next
    Dim visRegions As String: visRegions = ""
    Dim piR As PivotItem
    For Each piR In rfRegion.PivotItems
        visRegions = visRegions & piR.Value & "(vis=" & piR.Visible & ") "
    Next piR
    LogMsg "[RCD2] region items: " & left(visRegions, 200)

    ' Log visible center items
    Dim visCenters As String: visCenters = ""
    Dim piC As PivotItem
    For Each piC In rfCenter.PivotItems
        If piC.Visible Then visCenters = visCenters & piC.Value & " "
    Next piC
    LogMsg "[RCD2] visible centers: " & left(visCenters, 200)
    On Error GoTo SafeExit

    ' Write header
    pasteCell.Resize(1, 5).Value = Array("Region", "Center", "Total Budget", "Total Actual", "Ach %")

    ' Walk DataBodyRange rows directly - check cell content to detect hidden/filtered rows
    ' A filtered pivot row has blank values in all data cells
    Dim colRegion As Long, colCenter As Long, colBgt As Long, colAct As Long
    On Error Resume Next
    colRegion = rfRegion.DataRange.Column
    colCenter = rfCenter.DataRange.Column
    colBgt = dfBgt.DataRange.Column
    colAct = dfAct.DataRange.Column
    On Error GoTo SafeExit
    LogMsg "[RCD2] cols: region=" & colRegion & " center=" & colCenter & " bgt=" & colBgt & " act=" & colAct

    Dim dbrFirstRow As Long: dbrFirstRow = pt.DataBodyRange.Row
    Dim dbrRows     As Long: dbrRows = pt.DataBodyRange.Rows.Count
    Dim iOut As Long: iOut = 0
    Dim lastRegion As String: lastRegion = ""
    Dim r As Long

    On Error Resume Next
    For r = 0 To dbrRows - 1
        Dim centerVal As String: centerVal = ""
        centerVal = Trim$(CStr(wsPvt.Cells(dbrFirstRow + r, colCenter).Value))
        If Len(centerVal) = 0 Then GoTo nextRow

        ' Check row hidden
        Dim bHid As Boolean: bHid = wsPvt.Rows(dbrFirstRow + r).Hidden
        LogMsg "[RCD2] row=" & (dbrFirstRow + r) & " center=" & centerVal & " hidden=" & bHid
        If bHid Then GoTo nextRow

        Dim rv2 As String
        rv2 = Trim$(CStr(wsPvt.Cells(dbrFirstRow + r, colRegion).Value))
        If Len(rv2) > 0 Then lastRegion = rv2

        Dim dBgt As Double, dAct As Double
        dBgt = 0: dAct = 0
        dBgt = CDbl(wsPvt.Cells(dbrFirstRow + r, colBgt).Value)
        If Err.Number <> 0 Then dBgt = 0: Err.Clear
        dAct = CDbl(wsPvt.Cells(dbrFirstRow + r, colAct).Value)
        If Err.Number <> 0 Then dAct = 0: Err.Clear

        iOut = iOut + 1
        pasteCell.Offset(iOut, 0).Value = lastRegion
        pasteCell.Offset(iOut, 1).Value = centerVal
        pasteCell.Offset(iOut, 2).Value = dBgt
        pasteCell.Offset(iOut, 3).Value = dAct
        pasteCell.Offset(iOut, 4).Value = IIf(dBgt = 0, 0, dAct / dBgt)
nextRow:
    Next r
    On Error GoTo SafeExit

    LogMsg "[RCD2] iOut=" & iOut & " (rows written to AA)"

    pasteCell.Offset(1, 2).Resize(IIf(iOut > 0, iOut, 1), 1).NumberFormat = "#,##0"
    pasteCell.Offset(1, 3).Resize(IIf(iOut > 0, iOut, 1), 1).NumberFormat = "#,##0"
    pasteCell.Offset(1, 4).Resize(IIf(iOut > 0, iOut, 1), 1).NumberFormat = "0.0%"

    RefreshCountryChartData = iOut

SafeExit:
    LogMsg "[RCD2] SafeExit iOut=" & iOut & " err=" & Err.Number
End Function


Public Sub RepositionCountryDashboard(ByVal wsDash As Worksheet, ByVal country As String, Optional ByVal nPtsHint As Long = 0)
    ' Single sub: positions slicers AND chart for a country dashboard.
    ' Mirrors RepositionDashboardCharts strategy. Called only on build/ImportAndRefresh.
    On Error GoTo SafeExit
    Dim key As String: key = SanitizeName(country)

    ' Slicer sizing: H=4.87cm, W=4.53cm  (1cm = 28.35pt)
    Const SL_MARGIN As Double = 30   ' below row-1 title text
    Const SL_W      As Double = 128.4
    Const SL_H      As Double = 138.1
    Const SL_GAP    As Double = 10

    Dim slRegion As Shape, slMonth As Shape
    On Error Resume Next
    Set slRegion = wsDash.Shapes("slRegion_" & key)
    Set slMonth = wsDash.Shapes("slMonth_" & key)
    On Error GoTo SafeExit
    If Not slRegion Is Nothing Then
        slRegion.left = SL_MARGIN
        slRegion.top = SL_MARGIN
        slRegion.width = SL_W
        slRegion.height = SL_H
    End If
    If Not slMonth Is Nothing Then
        slMonth.left = SL_MARGIN + SL_W + SL_GAP
        slMonth.top = SL_MARGIN
        slMonth.width = SL_W
        slMonth.height = SL_H
    End If

    ' Charts: top = bottom of slicers + margin
    Const CH_MARGIN As Double = 6
    Const CH_GAP As Double = 8
    Dim chTop As Double: chTop = SL_MARGIN + SL_H + CH_MARGIN

    ' Base width: at least to rightmost slicer edge, but not less than worksheet viewport width.
    Dim totalW As Double: totalW = 0
    Dim shpScan As Shape
    For Each shpScan In wsDash.Shapes
        If shpScan.Type = 24 Then
            If shpScan.left + shpScan.width > totalW Then
                totalW = shpScan.left + shpScan.width
            End If
        End If
    Next shpScan
    totalW = Application.WorksheetFunction.Max(totalW, wsDash.Range("A1:V1").left + wsDash.Range("A1:V1").width)

    Dim chLeft As Double: chLeft = wsDash.Range("A1").left + CH_MARGIN

    ' Chart height scales with data point count
    Dim nPts As Long: nPts = 20
    On Error Resume Next
    Dim shpCh As Shape
    Set shpCh = wsDash.Shapes("chCenter_" & key)
    Dim shpAch As Shape
    Set shpAch = wsDash.Shapes("chAchCenter_" & key)
    If Not shpCh Is Nothing Then
        If shpCh.HasChart Then
            nPts = shpCh.Chart.SeriesCollection(1).Points.Count
        End If
    End If
    On Error GoTo SafeExit
    ' Use hint if provided (more reliable than Points.Count which can be stale)
    If nPtsHint > 0 Then nPts = nPtsHint

    Dim totalChartW As Double
    totalChartW = Application.WorksheetFunction.Max(600, totalW - chLeft - CH_MARGIN)

    Dim chH As Double
    chH = Application.WorksheetFunction.Min(500, _
          Application.WorksheetFunction.Max(280, nPts * 12 + 100))

    On Error Resume Next
    If Not shpCh Is Nothing Then
        shpCh.left = chLeft
        shpCh.top = chTop
        If Not shpAch Is Nothing Then
            shpCh.width = Application.WorksheetFunction.Max(300, (totalChartW - CH_GAP) / 2)
        Else
            shpCh.width = totalChartW
        End If
        shpCh.height = chH
        If shpCh.HasChart Then
            Dim ch2 As Chart: Set ch2 = shpCh.Chart
            If ch2.SeriesCollection.Count >= 1 Then
                ch2.ChartGroups(1).GapWidth = 80
            End If
            ' Reset PlotArea to auto so Excel fills chart area properly
            ch2.PlotArea.Position = xlAutomatic
        End If
    End If
    If Not shpAch Is Nothing Then
        Dim achLeftPos As Double
        achLeftPos = chLeft + Application.WorksheetFunction.Max(300, (totalChartW - CH_GAP) / 2) + CH_GAP
        If Not shpCh Is Nothing Then achLeftPos = shpCh.left + shpCh.width + CH_GAP
        shpAch.left = achLeftPos
        shpAch.top = chTop
        shpAch.width = Application.WorksheetFunction.Max(300, (totalChartW - CH_GAP) / 2)
        shpAch.height = chH
    End If
    On Error GoTo SafeExit
SafeExit:
End Sub



Public Sub EnsureCountryDashboardCharts()
    ' Called from both build and refresh paths.
    ' For each country dashboard:
    '   - If chart missing or is a placeholder: rebuild from static range
    '   - If chart exists: refresh data + reposition
    On Error GoTo SafeExit
    Dim dashMap As Object
    Set dashMap = CreateObject("Scripting.Dictionary")
    dashMap.CompareMode = vbTextCompare

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If left$(ws.name, 10) = "Dashboard_" Then
            If Not dashMap.Exists(ws.name) Then dashMap.Add ws.name, True
        End If
    Next ws

    ' Ensure every country in tblData has a dashboard/pivots pair.
    On Error Resume Next
    Dim wsData As Worksheet
    Dim lo As ListObject
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set lo = wsData.ListObjects("tblData")
    On Error GoTo SafeExit
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Dim countries As Collection
            Set countries = GetUniqueCountries(lo)
            Dim country As Variant
            For Each country In countries
                Dim c As String: c = Trim$(CStr(country))
                If Len(c) > 0 Then
                    Dim dashName As String
                    dashName = "Dashboard_" & SanitizeSheetName(c)
                    If Not dashMap.Exists(dashName) Then
                        If RebuildCountryDashboardFromData(c) Then
                            dashMap(dashName) = True
                        End If
                    End If
                End If
            Next country
        End If
    End If

    Dim nm As Variant
    For Each nm In dashMap.keys
        Dim wsDash As Worksheet
        Dim countryName As String
        Set wsDash = ThisWorkbook.Worksheets(CStr(nm))
        countryName = Mid$(wsDash.name, 11)   ' strip "Dashboard_"
        EnsureOneCountryChart wsDash, countryName
    Next nm
SafeExit:
End Sub

Private Sub EnsureOneCountryChart(ByVal wsDash As Worksheet, ByVal country As String)
    On Error GoTo SafeExit
    Dim key As String
    key = SanitizeName(country)
    Dim chartName As String
    chartName = "chCenter_" & key
    Dim achChartName As String
    achChartName = "chAchCenter_" & key

    Dim wsPvt As Worksheet
    On Error Resume Next
    Set wsPvt = ThisWorkbook.Worksheets("Pivots_" & SanitizeSheetName(country))
    On Error GoTo SafeExit
    If wsPvt Is Nothing Then Exit Sub

    Dim pt As PivotTable
    On Error Resume Next
    Set pt = wsPvt.PivotTables("ptCenter_" & key)
    On Error GoTo SafeExit
    If pt Is Nothing Then
        If RebuildCountryDashboardFromData(country) Then Exit Sub
        Exit Sub
    End If

    If Not IsCountryPivotHealthy(wsPvt, pt, country) Then
        If RebuildCountryDashboardFromData(country) Then Exit Sub
    End If

    If Not IsCountryPivotBoundToCountryHost(pt, country) Then
        LogMsg "[EOC] ptCenter_" & key & " is not bound to country host; rebuilding " & country
        If RebuildCountryDashboardFromData(country) Then Exit Sub
    End If

    If Not IsCountryPivotSourceCurrent(pt, country) Then
        LogMsg "[EOC] ptCenter_" & key & " source rows stale; rebuilding " & country
        If RebuildCountryDashboardFromData(country) Then Exit Sub
    End If

    ' Force refresh for this country's pivots to pull latest values when row count is unchanged.
    On Error Resume Next
    pt.PivotCache.Refresh
    pt.RefreshTable
    Dim ptCCRefresh As PivotTable
    Set ptCCRefresh = wsPvt.PivotTables("ptCC_" & key)
    If Not ptCCRefresh Is Nothing Then
        ptCCRefresh.PivotCache.Refresh
        ptCCRefresh.RefreshTable
    End If
    On Error GoTo SafeExit

    ' Enforce Ach data field for every country pivot (same build contract as MY).
    EnsureAchievementDataField pt
    pt.RefreshTable
    DoEvents
    If Not HasAchDataField(pt) Then
        LogMsg "[EOC] missing Ach data field in ptCenter_" & key & ", rebuilding " & country
        If RebuildCountryDashboardFromData(country) Then Exit Sub
    End If

    ' Write AZ1 country tag early - needed for event handler
    On Error Resume Next
    wsPvt.Range("AZ1").Value = country
    On Error GoTo SafeExit

    ' Check if charts exist and are real charts
    Dim shp As Shape
    Dim shpAch As Shape
    Dim chartMissing As Boolean
    chartMissing = True
    Dim achChartMissing As Boolean
    achChartMissing = True
    On Error Resume Next
    Set shp = wsDash.Shapes(chartName)
    If Not shp Is Nothing Then
        If shp.HasChart Then chartMissing = False
    End If
    Set shpAch = wsDash.Shapes(achChartName)
    If Not shpAch Is Nothing Then
        If shpAch.HasChart Then achChartMissing = False
    End If
    On Error GoTo SafeExit

    ' Chart exists but has no Ach series -> treat as missing and rebuild.
    If Not achChartMissing Then
        On Error Resume Next
        If Not ChartHasAchievementSeries(shpAch.Chart) Then
            achChartMissing = True
        End If
        On Error GoTo SafeExit
    End If

    ' Recreate slicers if missing
    Dim slicersMissing As Boolean: slicersMissing = True
    On Error Resume Next
    Dim shpSlCheck As Shape
    Set shpSlCheck = wsDash.Shapes("slRegion_" & key)
    If Not shpSlCheck Is Nothing Then slicersMissing = False
    On Error GoTo SafeExit
    If slicersMissing Then
        Dim ptCC2 As PivotTable
        On Error Resume Next
        Set ptCC2 = wsPvt.PivotTables("ptCC_" & key)
        On Error GoTo SafeExit
        If Not ptCC2 Is Nothing Then
            AddCountrySlicers wsDash, pt, ptCC2, country
            ' Let pivot settle after slicer recreation
            DoEvents
            On Error Resume Next
            pt.RefreshTable
            DoEvents
            On Error GoTo SafeExit
        End If
    End If

    If chartMissing Then
        ' Delete placeholder if present, rebuild as PivotChart
        On Error Resume Next
        If Not shp Is Nothing Then shp.Delete
        On Error GoTo SafeExit

        Dim chLeft As Double: chLeft = wsDash.Range("A1").left + 6
        Dim chTop  As Double: chTop = 175
        Set shp = AddPivotChartClean(wsDash, wsPvt, pt, xlColumnClustered, _
                    "Center Performance (" & country & ")", _
                    chLeft, chTop, 800, 420, chartName)
        LogMsg "[EOC] PivotChart rebuilt for " & country & " err=" & Err.Number: Err.Clear
    End If

    If achChartMissing Then
        On Error Resume Next
        If Not shpAch Is Nothing Then shpAch.Delete
        On Error GoTo SafeExit

        EnsureAchievementDataField pt
        pt.RefreshTable
        DoEvents

        Dim achLeft As Double: achLeft = wsDash.Range("A1").left + 414
        Set shpAch = AddAchPivotChart(wsDash, wsPvt, pt, _
                    achLeft, 175, 390, 420, achChartName, _
                    "Achievement % (" & country & ")")
        LogMsg "[EOC] chAchCenter rebuilt for " & country & " err=" & Err.Number: Err.Clear
    End If

    ' Always enforce Ach-only axis formatting (fixes existing charts too).
    On Error Resume Next
    Set shpAch = wsDash.Shapes(achChartName)
    If Not shpAch Is Nothing Then
        If shpAch.HasChart Then ApplyAchOnlySeriesFormat shpAch.Chart
    End If
    On Error GoTo SafeExit

    ' Always reposition slicers + chart
    RepositionCountryDashboard wsDash, country

SafeExit:
End Sub


Private Function IsCountryPivotHealthy(ByVal wsPvt As Worksheet, ByVal pt As PivotTable, ByVal country As String) As Boolean
    On Error GoTo Bad
    If wsPvt Is Nothing Then GoTo Bad
    If pt Is Nothing Then GoTo Bad

    ' Detect #SPILL! or similar block errors near pivot anchor.
    Dim chk As Range
    Set chk = wsPvt.Range("A1:H20")
    Dim c As Range
    For Each c In chk.Cells
        If InStr(1, CStr(c.Text), "#SPILL!", vbTextCompare) > 0 Then
            LogMsg "[EOC] unhealthy pivot (" & country & "): spill marker at " & c.Address(False, False)
            GoTo Bad
        End If
    Next c

    If pt.TableRange2 Is Nothing Then GoTo Bad
    If pt.DataFields.Count < 2 Then GoTo Bad

    IsCountryPivotHealthy = True
    Exit Function
Bad:
    IsCountryPivotHealthy = False
End Function

Private Function IsCountryPivotBoundToCountryHost(ByVal pt As PivotTable, ByVal country As String) As Boolean
    On Error GoTo Bad
    If pt Is Nothing Then GoTo Bad

    Dim expectedHost As String
    expectedHost = BuildCountryCacheHostSheetName(country)

    Dim src As String
    src = CStr(pt.PivotCache.SourceData)
    If InStr(1, src, expectedHost, vbTextCompare) <= 0 Then GoTo Bad

    IsCountryPivotBoundToCountryHost = True
    Exit Function
Bad:
    IsCountryPivotBoundToCountryHost = False
End Function

Private Function CountCountrySourceRowsFromTblData(ByVal country As String) As Long
    ' Returns expected source rows for country cache host:
    ' header row + data rows matching the same filter used by RestageOneCountryCacheHost.
    On Error GoTo SafeExit

    Dim wsData As Worksheet
    Dim lo As ListObject
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set lo = wsData.ListObjects("tblData")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim key As String
    key = NormalizeCountryKey(country)

    Dim cnt As Long
    cnt = 0

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim rowKey As String
        Dim regionVal As String
        Dim centerVal As String
        Dim monthVal As String

        rowKey = NormalizeCountryKey(lo.DataBodyRange.Cells(r, 1).Value2)
        regionVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        centerVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2))
        monthVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))

        If rowKey = key _
           And Len(regionVal) > 0 _
           And Len(centerVal) > 0 _
           And Len(monthVal) > 0 Then
            cnt = cnt + 1
        End If
    Next r

    If cnt > 0 Then
        CountCountrySourceRowsFromTblData = cnt + 1 ' + header row
    End If
SafeExit:
End Function

Private Function IsCountryPivotSourceCurrent(ByVal pt As PivotTable, ByVal country As String) As Boolean
    On Error GoTo Bad
    If pt Is Nothing Then GoTo Bad

    Dim expectedRows As Long
    expectedRows = CountCountrySourceRowsFromTblData(country)
    If expectedRows <= 1 Then GoTo Bad

    Dim srcRows As Long
    srcRows = ExtractSourceLastRow(CStr(pt.PivotCache.SourceData))
    If srcRows <> expectedRows Then GoTo Bad

    IsCountryPivotSourceCurrent = True
    Exit Function
Bad:
    IsCountryPivotSourceCurrent = False
End Function

Private Function RebuildCountryDashboardFromData(ByVal country As String) As Boolean
    On Error GoTo SafeExit
    country = NormalizeCountryKey(country)
    Dim wsData As Worksheet
    Dim lo As ListObject
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set lo = wsData.ListObjects("tblData")
    If lo Is Nothing Then GoTo SafeExit
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit

    LogMsg "[EOC] rebuilding country dashboard from data: " & country
    BuildOneCountryDashboard country, lo
    RebuildCountryDashboardFromData = True
    Exit Function
SafeExit:
End Function


Private Sub ApplyAchOnlySeriesFormat(ByVal ch As Chart)
    On Error Resume Next

    Dim s As Series
    Dim achSeriesCount As Long
    achSeriesCount = 0

    For Each s In ch.SeriesCollection
        Dim nm As String: nm = LCase$(CStr(s.name))
        If InStr(nm, "ctrach") > 0 Or InStr(nm, "ach") > 0 Or InStr(nm, "%") > 0 Then
            achSeriesCount = achSeriesCount + 1
            s.chartType = xlLineMarkers
            s.AxisGroup = xlPrimary
            s.HasDataLabels = True
            s.DataLabels.ShowValue = True
            s.DataLabels.NumberFormat = "0.0%"
            s.DataLabels.Position = xlLabelPositionAbove
            s.DataLabels.Font.size = 9
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 6
        Else
            ' Keep hidden non-Ach series off the primary axis so Y scale stays in %.
            s.AxisGroup = xlSecondary
            s.Format.Fill.Visible = msoFalse
            s.Format.Line.Visible = msoFalse
            s.HasDataLabels = False
        End If
        Err.Clear
    Next s

    ch.HasTitle = True
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    ch.ShowLegendFieldButtons = False
    ch.ShowAxisFieldButtons = False
    ch.ShowValueFieldButtons = False
    ch.ShowReportFilterFieldButtons = False

    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0%"
    ch.Axes(xlValue, xlPrimary).MinimumScale = 0
    ch.Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True

    If ch.HasAxis(xlValue, xlSecondary) Then
        ch.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
        ch.Axes(xlValue, xlSecondary).Format.Line.Visible = msoFalse
    End If

    LogMsg "[AAX] formatted Ach chart series=" & ch.SeriesCollection.Count & " achSeries=" & achSeriesCount
End Sub


Private Function HasAchDataField(ByVal pt As PivotTable) As Boolean
    On Error GoTo SafeExit
    Dim i As Long
    For i = 1 To pt.DataFields.Count
        Dim nm As String
        nm = LCase$(CStr(pt.DataFields(i).name))
        If InStr(nm, "ctrach") > 0 _
           Or InStr(nm, "ach %") > 0 _
           Or InStr(nm, "achievementpct") > 0 _
           Or InStr(nm, "achievement %") > 0 Then
            HasAchDataField = True
            Exit Function
        End If
    Next i
SafeExit:
End Function

Private Function ChartHasAchievementSeries(ByVal ch As Chart) As Boolean
    On Error GoTo SafeExit
    Dim s As Series
    For Each s In ch.SeriesCollection
        Dim nm As String
        nm = LCase$(CStr(s.name))
        If InStr(nm, "ctrach") > 0 _
           Or InStr(nm, "ach %") > 0 _
           Or InStr(nm, "achievement") > 0 Then
            ChartHasAchievementSeries = True
            Exit Function
        End If
    Next s
SafeExit:
End Function

Private Sub EnsureAchievementDataField(ByVal pt As PivotTable)
    On Error GoTo SafeExit
    If pt Is Nothing Then Exit Sub
    pt.ManualUpdate = True

    ' Remove old Ach-related data fields on this pivot only.
    Dim i As Long
    For i = pt.DataFields.Count To 1 Step -1
        Dim dfNm As String
        dfNm = LCase$(CStr(pt.DataFields(i).name))
        If InStr(dfNm, "ctrach") > 0 _
           Or InStr(dfNm, "ach %") > 0 _
           Or InStr(dfNm, "achievement") > 0 _
           Or InStr(dfNm, "achievementpct") > 0 Then
            pt.DataFields(i).Orientation = xlHidden
        End If
    Next i

    ' Ensure calculated field exists; do not delete global calculated fields
    ' because they may be in use by other pivots.
    Dim hasCtrAch As Boolean
    hasCtrAch = False
    Dim cf As Object
    On Error Resume Next
    For Each cf In pt.CalculatedFields
        If LCase$(CStr(cf.name)) = "ctrach" Then
            hasCtrAch = True
            Exit For
        End If
    Next cf
    On Error GoTo SafeExit
    If Not hasCtrAch Then
        On Error Resume Next
        pt.CalculatedFields.Add "CtrAch", "=Actual/Budget", True
        Err.Clear
        On Error GoTo SafeExit
    End If

    Dim pfCtr As PivotField
    On Error Resume Next
    Set pfCtr = pt.PivotFields("CtrAch")
    On Error GoTo SafeExit
    If pfCtr Is Nothing Then GoTo SafeExit

    Dim dfAch As PivotField
    Set dfAch = pt.AddDataField(pfCtr, "Ach %", xlSum)
    If Not dfAch Is Nothing Then dfAch.NumberFormat = "0.0%"

    pt.ManualUpdate = False
    pt.RefreshTable

SafeExit:
    On Error Resume Next
    pt.ManualUpdate = False
    On Error GoTo 0
End Sub



Public Sub LogMsg(ByVal msg As String)
    On Error Resume Next
    Dim fNum As Integer: fNum = FreeFile
    Open "C:\temp\dashboard_log.txt" For Append As #fNum
    Print #fNum, Format(Now, "hh:mm:ss") & "  " & msg
    Close #fNum
    On Error GoTo 0
End Sub

Public Sub ClearLog()
    On Error Resume Next
    Kill "C:\temp\dashboard_log.txt"
    MkDir "C:\temp"
    On Error GoTo 0
End Sub


Public Sub RunPendingChartRefresh()
    ' Legacy compatibility callback.
    ' Old worksheet event code may still schedule:
    '   Application.OnTime ..., "Main.RunPendingChartRefresh"
    ' Keep this entry point to avoid runtime macro-not-found errors.
    On Error GoTo SafeExit

    Dim country As String
    country = Trim$(CStr(gPendingChartCountry))
    gPendingChartCountry = ""

    If Len(country) > 0 Then
        LogMsg "[RPC] requested country=" & country
        If Not RebuildCountryDashboardFromData(country) Then
            Dim wsDash As Worksheet
            On Error Resume Next
            Set wsDash = ThisWorkbook.Worksheets("Dashboard_" & SanitizeSheetName(country))
            On Error GoTo SafeExit
            If Not wsDash Is Nothing Then
                EnsureOneCountryChart wsDash, country
            Else
                EnsureCountryDashboardCharts
            End If
        End If
    Else
        LogMsg "[RPC] no pending country; ensure all"
        EnsureCountryDashboardCharts
    End If

SafeExit:
End Sub


Public Sub RepairCountryDashboard(Optional ByVal country As String = "SG")
    On Error GoTo EH
    Dim c As String
    c = Trim$(country)
    If Len(c) = 0 Then Exit Sub

    If Not RebuildCountryDashboardFromData(c) Then
        Err.Raise vbObjectError + 7901, , "Unable to rebuild country dashboard for [" & c & "]."
    End If
    EnsureCountryDashboardCharts
    Exit Sub
EH:
    MsgBox "RepairCountryDashboard failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub


Public Sub RunPendingAxisRebuild()
    ' Strategy 11 Phase 2: runs 3 seconds after slicer click.
    ' Slicer lock fully released by now - SERIES formula assignment works.
    If Len(gPendingAxisCountry) = 0 Then Exit Sub
    Dim c As String: c = gPendingAxisCountry
    gPendingAxisCountry = ""
    On Error GoTo SafeExit

    Dim key As String: key = SanitizeName(c)
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard_" & SanitizeSheetName(c))
    On Error GoTo SafeExit
    If wsDash Is Nothing Then GoTo SafeExit

    Dim srcRng As Range
    Set srcRng = wsDash.Range("AA1").CurrentRegion
    If srcRng.Rows.Count < 2 Then GoTo SafeExit

    Dim nRows As Long: nRows = srcRng.Rows.Count - 1
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsDash.Shapes("chCenter_" & key)
    On Error GoTo SafeExit
    If shp Is Nothing Then GoTo SafeExit

    Dim ch As Chart: Set ch = shp.Chart
    Dim shn As String: shn = "'" & wsDash.name & "'"
    Dim catA As String: catA = shn & "!" & srcRng.Columns(2).Offset(1, 0).Resize(nRows).Address(True, True)
    Dim bgtA As String: bgtA = shn & "!" & srcRng.Columns(3).Offset(1, 0).Resize(nRows).Address(True, True)
    Dim actA As String: actA = shn & "!" & srcRng.Columns(4).Offset(1, 0).Resize(nRows).Address(True, True)
    Dim achA As String: achA = shn & "!" & srcRng.Columns(5).Offset(1, 0).Resize(nRows).Address(True, True)

    On Error Resume Next
    Dim sc As Long: sc = ch.SeriesCollection.Count
    If sc >= 1 Then
        ch.SeriesCollection(1).Formula = "=SERIES(""Total Budget""," & catA & "," & bgtA & ",1)"
        LogMsg "[P2] s1 formula err=" & Err.Number: Err.Clear
    End If
    If sc >= 2 Then
        ch.SeriesCollection(2).Formula = "=SERIES(""Total Actual""," & catA & "," & actA & ",2)"
        LogMsg "[P2] s2 formula err=" & Err.Number: Err.Clear
    End If
    If sc >= 3 Then
        ch.SeriesCollection(3).Formula = "=SERIES(""Ach %""," & catA & "," & achA & ",3)"
        LogMsg "[P2] s3 formula err=" & Err.Number: Err.Clear
    End If

    ' Reposition chart width to match new data point count
    RepositionCountryDashboard wsDash, c, nRows
    LogMsg "[P2] complete country=" & c
SafeExit:
End Sub





'----------------------------------------------------------------
' Build one country dashboard + pivots from tblData.
'----------------------------------------------------------------
Private Sub BuildOneCountryDashboard(ByVal country As String, ByVal lo As ListObject)
    country = NormalizeCountryKey(country)
    gStep = "BuildOneCountryDashboard: " & country
    On Error GoTo EH

    '--------------------------------------------
    ' Sheet names
    '--------------------------------------------
    gStep = "BuildOneCountryDashboard: sanitizing sheet name " & country

    Dim dashName As String, pvtName As String
    dashName = "Dashboard_" & SanitizeSheetName(country)
    pvtName = "Pivots_" & SanitizeSheetName(country)

    Dim wsDash As Worksheet, wsPvt As Worksheet
    Set wsDash = CreateOrResetSheet(dashName, ThisWorkbook.Worksheets("Dashboard"))
    Set wsPvt = CreateOrResetSheet(pvtName, wsDash)

    '--------------------------------------------
    ' Create COUNTRY-ONLY PivotCache (THIS FIXES REGION LIST)
    '--------------------------------------------
    gStep = "BuildOneCountryDashboard: create country PivotCache " & country
    Dim pc As PivotCache
    Set pc = CreatePivotCacheFromTable_FilterCountry(lo, country)
    pc.MissingItemsLimit = xlMissingItemsNone

    '--------------------------------------------
    ' Detach slicers referencing pivots on this Pivots_<country> sheet (safety)
    '--------------------------------------------
    gStep = "BuildOneCountryDashboard: detach slicer referencing " & country
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        Dim i As Long
        For i = sc.PivotTables.Count To 1 Step -1
            Dim p As PivotTable
            Set p = sc.PivotTables(i)
            If p.Parent.name = wsPvt.name Then
                sc.PivotTables.RemovePivotTable p
            End If
        Next i
    Next sc

    '--------------------------------------------
    ' Ensure the pivots sheet is truly empty of pivots
    '--------------------------------------------
    gStep = "BuildOneCountryDashboard: clear pivots sheet " & country
    ClearPivotTablesOnly wsPvt
    wsPvt.Cells.Clear
    wsPvt.Cells.ClearFormats
    wsPvt.Cells.UnMerge

    Dim key As String
    key = SanitizeName(country)

    Dim ptCenter As PivotTable, ptCC As PivotTable

    '==================================================
    ' Pivot 1: ptCenter_<country>  (Row: Region, Center; Filters: Country, Month)
    '==================================================
    gStep = "BuildOneCountryDashboard: create ptCenter " & country

    DeletePivotIfExists wsPvt, "ptCenter_" & key
    Set ptCenter = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Range("A1"), _
        TableName:="ptCenter_" & key)

    With ptCenter
        .ManualUpdate = True
        .ClearAllFilters

        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Center").Orientation = xlRowField

        .PivotFields("Country").Orientation = xlPageField
        .PivotFields("Month").Orientation = xlPageField

        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum

        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels

        .ManualUpdate = False
    End With

    ptCenter.RefreshTable
    EnsureAchievementDataField ptCenter
    ForceMonthDisplay ptCenter
    DoEvents

    ' Force page filter to this country (should always exist because cache is already filtered)
    SafeSetPageFilter ptCenter, "Country", country
    ptCenter.RefreshTable
    DoEvents

    ' Ensure Region/Center dropdown only show items with data
    RestrictRegionDropdown ptCenter

    On Error GoTo 0

    '==================================================
    ' Pivot 2: ptCC_<country> (Row: Region, Center; Filter: Country)
    '==================================================
    Dim nextRow As Long
    nextRow = ptCenter.TableRange2.Row + ptCenter.TableRange2.Rows.Count + 3

    gStep = "BuildOneCountryDashboard: create ptCC " & country

    DeletePivotIfExists wsPvt, "ptCC_" & key
    Set ptCC = pc.CreatePivotTable( _
        TableDestination:=wsPvt.Cells(nextRow, 1), _
        TableName:="ptCC_" & key)

    With ptCC
        .ManualUpdate = True
        .ClearAllFilters

        .PivotFields("Country").Orientation = xlPageField
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Center").Orientation = xlRowField

        .AddDataField .PivotFields("Budget"), "Total Budget", xlSum
        .AddDataField .PivotFields("Actual"), "Total Actual", xlSum

        .RowAxisLayout xlTabularRow

        .ManualUpdate = False
    End With

    ptCC.RefreshTable
    DoEvents

    SafeSetPageFilter ptCC, "Country", country
    ptCC.RefreshTable
    DoEvents

    RestrictRegionDropdown ptCC

    '==================================================
    ' Dashboard visuals
    ' IMPORTANT: Clear dashboard BEFORE adding slicers
    '==================================================
    gStep = "BuildOneCountryDashboard: clear dashboard area " & country
    wsDash.Cells.ClearContents
    wsDash.Cells.ClearFormats

    wsDash.Range("A1").Value = "Country Dashboard: " & country
    wsDash.Range("A1").Font.Bold = True
    wsDash.Range("A1").Font.size = 16

    ' Country slicers (Region + Month) for this country dashboard
    gStep = "BuildOneCountryDashboard: add slicers " & country
    AddCountrySlicers wsDash, ptCenter, ptCC, country
    RepositionCountryDashboard wsDash, country

    ' Two side-by-side charts below the slicers:
    '   Left  (chCenter_MY)    : Budget + Actual clustered columns  - PivotChart, auto-updates
    '   Right (chAchCenter_MY) : Achievement % line                 - PivotChart, auto-updates
    Const CH_TOP  As Double = 175
    Const CH_H    As Double = 380
    Const CH_GAP  As Double = 8
    Dim totalW  As Double: totalW = wsDash.Range("A1").left + 900    ' ~col N
    Dim chLeft  As Double: chLeft = wsDash.Range("A1").left + 6
    Dim chW     As Double: chW = (totalW - chLeft - CH_GAP) / 2       ' half width each

    ' Left chart: PivotChart - Budget vs Actual columns
    gStep = "BuildOneCountryDashboard: chCenter " & country
    Dim shp As Shape
    Set shp = AddPivotChartClean(wsDash, wsPvt, ptCenter, xlColumnClustered, _
                    "Budget vs Actual (" & country & ")", _
                    chLeft, CH_TOP, chW, CH_H, "chCenter_" & key)
    LogMsg "[BOC] chCenter built err=" & Err.Number: Err.Clear

    ' Right chart: second PivotChart bound to same ptCenter — shows CtrAch line only.
    ' Because it shares ptCenter, slicer clicks update BOTH charts automatically.
    ' Zero VBA on slicer click. Zero lock issues.
    gStep = "BuildOneCountryDashboard: chAchCenter " & country
    Dim achLeft As Double: achLeft = chLeft + chW + CH_GAP
    Set shp = AddAchPivotChart(wsDash, wsPvt, ptCenter, _
                    achLeft, CH_TOP, chW, CH_H, _
                    "chAchCenter_" & key, _
                    "Achievement % (" & country & ")")
    LogMsg "[BOC] chAchCenter built err=" & Err.Number: Err.Clear

    ' Position: full width below slicers
    RepositionCountryDashboard wsDash, country

    ' Breakdown: Country x Center
    gStep = "BuildOneCountryDashboard: breakdown Country x Center " & country
    RenderCountryCenterTable wsDash, ptCC, ptCC.TableRange1.Cells(1, 1)

    On Error Resume Next
    FormatDashboard wsDash
    On Error GoTo 0

    gStep = "BuildOneCountryDashboard: done " & country
    Exit Sub

EH:
    Err.Raise vbObjectError + 7703, , gStep & ": " & Err.Number & " - " & Err.Description
End Sub

'----------------------------------------------------------------
' Writes Ach% helper formulas to wsPvt cols AY:AZ (cols 51-52).
' AY = Center label,  AZ = =Actual/Budget (live formula)
' These reference ptCenter DataBodyRange cells directly so they
' recalculate automatically when the pivot refreshes on slicer click.
'----------------------------------------------------------------


'----------------------------------------------------------------
' Builds a second PivotChart bound to ptCenter showing ONLY the
' CtrAch series as a line chart. Shares the same pivot as chCenter
' so slicer clicks update both charts simultaneously — zero VBA needed.
'----------------------------------------------------------------
Private Function AddAchPivotChart( _
    ByVal wsDash As Worksheet, ByVal wsPvt As Worksheet, _
    ByVal pt As PivotTable, _
    ByVal chLeft As Double, ByVal chTop As Double, _
    ByVal chW As Double, ByVal chH As Double, _
    ByVal chartName As String, ByVal title As String) As Shape

    On Error Resume Next
    wsDash.Shapes(chartName).Delete: Err.Clear
    On Error GoTo 0

    ' Create chart shape (then bind to pivot table)
    Dim shp As Shape
    Set shp = wsDash.Shapes.AddChart(xlColumnClustered, chLeft, chTop, chW, chH)
    If shp Is Nothing Then
        LogMsg "[AAP] AddChart failed for " & chartName
        Set AddAchPivotChart = Nothing
        Exit Function
    End If
    shp.name = chartName

    Dim ch As Chart
    Set ch = shp.Chart

    ' Bind to ptCenter as a PivotChart.
    ' Some Excel builds fail on PivotLayout assignment for fresh charts.
    Dim bindErr As Long
    On Error Resume Next
    ch.PivotLayout.PivotTable = pt
    bindErr = Err.Number: Err.Clear
    If bindErr <> 0 Then
        ch.SetSourceData Source:=pt.TableRange2
        bindErr = Err.Number: Err.Clear
    End If
    LogMsg "[AAP] PivotLayout bind err=" & bindErr
    On Error GoTo 0

    ch.HasTitle = True
    ch.chartTitle.Text = title
    ApplyAchOnlySeriesFormat ch
    Err.Clear

    LogMsg "[AAP] built " & chartName & " series=" & ch.SeriesCollection.Count
    Set AddAchPivotChart = shp
End Function

'----------------------------------------------------------------
' Writes Ach% helper formulas to wsPvt cols AY:AZ (cols 51:52).
'   AY = Center label  (value, copied from pivot DataBodyRange)
'   AZ = =Actual/Budget (live formula referencing pivot cells)
' These recalculate automatically when ptCenter refreshes on slicer.
' Called once at build time. Safe to re-call on reimport.
'----------------------------------------------------------------
Private Sub WriteAchToPvtSheet(ByVal wsPvt As Worksheet, ByVal pt As PivotTable)
    On Error Resume Next
    Const LBL_COL As Long = 51  ' AY
    Const ACH_COL As Long = 52  ' AZ

    wsPvt.Cells(1, LBL_COL).Resize(300, 2).ClearContents

    If pt.DataBodyRange Is Nothing Then
        LogMsg "[WAT] DataBodyRange is Nothing, skip"
        Exit Sub
    End If

    Dim colCtr As Long: colCtr = pt.PivotFields("Center").DataRange.Column
    Dim colBgt As Long: colBgt = pt.PivotFields("Total Budget").DataRange.Column
    Dim colAct As Long: colAct = pt.PivotFields("Total Actual").DataRange.Column

    Dim firstRow As Long: firstRow = pt.DataBodyRange.Row
    Dim nRows    As Long: nRows = pt.DataBodyRange.Rows.Count

    wsPvt.Cells(1, LBL_COL).Value = "Center"
    wsPvt.Cells(1, ACH_COL).Value = "Ach %"

    Dim iOut As Long: iOut = 0
    Dim r    As Long
    For r = 0 To nRows - 1
        Dim cVal As String
        cVal = Trim$(CStr(wsPvt.Cells(firstRow + r, colCtr).Value))
        If Len(cVal) = 0 Then GoTo NextR

        iOut = iOut + 1
        wsPvt.Cells(1 + iOut, LBL_COL).Value = cVal
        Dim bA As String: bA = wsPvt.Cells(firstRow + r, colBgt).Address(True, True)
        Dim aA As String: aA = wsPvt.Cells(firstRow + r, colAct).Address(True, True)
        wsPvt.Cells(1 + iOut, ACH_COL).Formula = "=IFERROR(" & aA & "/" & bA & ",0)"
        wsPvt.Cells(1 + iOut, ACH_COL).NumberFormat = "0.0%"
NextR:
    Next r
    LogMsg "[WAT] wrote " & iOut & " rows to " & wsPvt.name & "!AY:AZ err=" & Err.Number
End Sub

'----------------------------------------------------------------
' Builds a plain static line chart on wsDash reading Ach% formula
' cells from wsPvt AY:AZ.
' - Category labels : wsPvt!AY2:AY(n+1)
' - Values          : wsPvt!AZ2:AZ(n+1)  (live =Actual/Budget formulas)
' Chart auto-recalculates when pivot refreshes. Zero VBA on slicer.
'----------------------------------------------------------------
Private Sub BuildAchLineChart( _
    ByVal wsDash As Worksheet, ByVal wsPvt As Worksheet, _
    ByVal chLeft As Double, ByVal chTop As Double, _
    ByVal chW As Double, ByVal chH As Double, _
    ByVal chartName As String, ByVal title As String)

    On Error Resume Next
    Const LBL_COL As Long = 51  ' AY
    Const ACH_COL As Long = 52  ' AZ

    ' Delete old shape with same name
    wsDash.Shapes(chartName).Delete: Err.Clear

    ' Count rows written by WriteAchToPvtSheet
    Dim nRows As Long: nRows = 0
    Dim r As Long
    For r = 1 To 299
        If Len(Trim$(CStr(wsPvt.Cells(r + 1, LBL_COL).Value))) = 0 Then Exit For
        nRows = r
    Next r

    If nRows < 1 Then
        LogMsg "[BAL] no Ach rows on " & wsPvt.name & ", skip chart"
        Exit Sub
    End If

    ' SERIES formula addresses — reference wsPvt cells directly
    Dim pvtShn As String: pvtShn = "'" & wsPvt.name & "'"
    Dim catA As String: catA = pvtShn & "!" & wsPvt.Cells(2, LBL_COL).Resize(nRows).Address(True, True)
    Dim achA As String: achA = pvtShn & "!" & wsPvt.Cells(2, ACH_COL).Resize(nRows).Address(True, True)

    ' Create plain chart (NOT PivotChart — no lock issues)
    Dim shp As Shape
    Set shp = wsDash.Shapes.AddChart(xlLine, chLeft, chTop, chW, chH)
    If shp Is Nothing Then LogMsg "[BAL] AddChart failed": Exit Sub
    shp.name = chartName

    Dim ch As Chart: Set ch = shp.Chart

    ' Clear auto-created series
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete: Err.Clear
    Loop

    ' Seed + add series (NewSeries needs at least one existing series or SetSourceData first)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    If s Is Nothing Then
        Application.EnableEvents = False
        ch.SetSourceData Source:=wsPvt.Cells(1, ACH_COL).Resize(1, 1)
        Application.EnableEvents = True
        Err.Clear
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete: Err.Clear
        Loop
        Set s = ch.SeriesCollection.NewSeries: Err.Clear
    End If
    If s Is Nothing Then LogMsg "[BAL] NewSeries failed": Exit Sub

    ' Pin SERIES formula to wsPvt formula cells
    s.Formula = "=SERIES(""Ach %""," & catA & "," & achA & ",1)"
    LogMsg "[BAL] SERIES formula err=" & Err.Number & " formula=" & s.Formula: Err.Clear

    ' Formatting
    ch.HasTitle = True
    ch.chartTitle.Text = title
    ch.HasLegend = False
    With s
        .chartType = xlLineMarkers
        .HasDataLabels = True
        .DataLabels.ShowValue = True
        .DataLabels.NumberFormat = "0.0%"
        .DataLabels.Position = xlLabelPositionAbove
        .DataLabels.Font.size = 9
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 6
    End With
    Err.Clear
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    ch.Axes(xlValue).MinimumScale = 0
    ch.PlotArea.Interior.Color = RGB(250, 250, 250)
    Err.Clear
    LogMsg "[BAL] built " & chartName & " nRows=" & nRows
End Sub

'==================================================
' DEBUG: Call this manually to trace chart build failure
' Run: DebugCountryChart "CN"
'==================================================
Public Sub DebugCountryChart(Optional ByVal country As String = "CN")
    Dim key As String: key = SanitizeName(country)
    Dim wsPvt As Worksheet, wsDash As Worksheet
    Dim pt As PivotTable
    Dim msg As String

    On Error Resume Next
    Set wsPvt = ThisWorkbook.Worksheets("Pivots_" & SanitizeSheetName(country))
    Set wsDash = ThisWorkbook.Worksheets("Dashboard_" & SanitizeSheetName(country))
    On Error GoTo 0

    msg = "=== DebugCountryChart: " & country & " ===" & vbCrLf
    msg = msg & "wsPvt: " & IIf(wsPvt Is Nothing, "MISSING", wsPvt.name) & vbCrLf
    msg = msg & "wsDash: " & IIf(wsDash Is Nothing, "MISSING", wsDash.name) & vbCrLf

    If wsPvt Is Nothing Or wsDash Is Nothing Then MsgBox msg: Exit Sub

    ' Pivot
    On Error Resume Next
    Set pt = wsPvt.PivotTables("ptCenter_" & key)
    On Error GoTo 0
    msg = msg & "ptCenter: " & IIf(pt Is Nothing, "MISSING", pt.name) & vbCrLf

    If Not pt Is Nothing Then
        msg = msg & "DataBodyRange: " & IIf(pt.DataBodyRange Is Nothing, "NOTHING", pt.DataBodyRange.Address) & vbCrLf
        msg = msg & "RowFields count: " & pt.RowFields.Count & vbCrLf
        Dim i As Long
        For i = 1 To pt.RowFields.Count
            msg = msg & "  RowField(" & i & "): " & pt.RowFields(i).name & vbCrLf
        Next i
        msg = msg & "DataFields count: " & pt.DataFields.Count & vbCrLf
        For i = 1 To pt.DataFields.Count
            msg = msg & "  DataField(" & i & "): " & pt.DataFields(i).name & vbCrLf
        Next i
    End If

    ' AA1 range
    Dim aa1 As Range
    Set aa1 = wsPvt.Range("AA1")
    Dim cr As Range
    Set cr = aa1.CurrentRegion
    msg = msg & "AA1 value: [" & CStr(aa1.Value) & "]" & vbCrLf
    msg = msg & "AA1.CurrentRegion rows: " & cr.Rows.Count & vbCrLf
    If cr.Rows.Count > 1 Then
        msg = msg & "AA1 data row1: " & cr.Cells(2, 1).Value & " | " & cr.Cells(2, 2).Value & " | " & cr.Cells(2, 3).Value & vbCrLf
    End If

    ' Shapes on wsDash + series names
    msg = msg & "wsDash shapes: " & wsDash.Shapes.Count & vbCrLf
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        msg = msg & "  [" & shp.name & "] HasChart=" & shp.HasChart & vbCrLf
        If shp.HasChart Then
            Dim s As Series
            Dim si As Long: si = 0
            For Each s In shp.Chart.SeriesCollection
                si = si + 1
                msg = msg & "    Series(" & si & ") name=[" & s.name & "]" & vbCrLf
            Next s
        End If
    Next shp

    ' AZ1 (country tag)
    msg = msg & "AZ1: [" & CStr(wsPvt.Range("AZ1").Value) & "]" & vbCrLf

    MsgBox msg, vbInformation, "DebugCountryChart"

    ' Now try to write data and report result
    If Not pt Is Nothing Then
        RefreshCountryChartData wsPvt, wsDash, pt, True
        Set cr = wsPvt.Range("AA1").CurrentRegion
        Dim msg2 As String
        msg2 = "After RefreshCountryChartData:" & vbCrLf
        msg2 = msg2 & "AA1.CurrentRegion rows: " & cr.Rows.Count & vbCrLf
        If cr.Rows.Count > 1 Then
            Dim r2 As Long
            For r2 = 2 To Application.WorksheetFunction.Min(cr.Rows.Count, 6)
                msg2 = msg2 & "  " & cr.Cells(r2, 1).Value & " | " & cr.Cells(r2, 2).Value & " | " & cr.Cells(r2, 3).Value & " | " & cr.Cells(r2, 4).Value & vbCrLf
            Next r2
        End If
        MsgBox msg2, vbInformation, "DebugCountryChart - Data"
    End If
End Sub


'==================================================
' Helpers
'==================================================

Private Function NormalizeCountryKey(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeCountryKey = UCase$(Trim$(s))
End Function

Private Function GetUniqueCountries(ByVal lo As ListObject) As Collection
    Dim col As New Collection
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim v As String
        v = NormalizeCountryKey(lo.DataBodyRange.Cells(r, 1).Value2) ' Country = col 1
        If Len(v) > 0 Then
            If Not dict.Exists(v) Then
                dict.Add v, True
                col.Add v
            End If
        End If
    Next r

    Set GetUniqueCountries = col
End Function

Private Function CreateOrResetSheet(ByVal sheetName As String, ByVal afterSheet As Worksheet) As Worksheet
    gStep = "CreateOrResetSheet: start [" & sheetName & "]"

    Dim ws As Worksheet
    On Error GoTo EH

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo EH

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        ws.name = sheetName
        Set CreateOrResetSheet = ws
        Exit Function
    End If

    ' Always show it while rebuilding
    ws.Visible = xlSheetVisible

    ' IMPORTANT:
    ' If the sheet ever hosted PivotTables, DO NOT ClearContents/ClearFormats.
    ' Delete & recreate the sheet. This avoids 1004 ClearContents failures.
    If ws.PivotTables.Count > 0 Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True

        Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        ws.name = sheetName
        Set CreateOrResetSheet = ws
        Exit Function
    End If

    ' No pivots -> safe to clear
    ws.Cells.Clear
    ws.Cells.ClearFormats
    ws.Cells.UnMerge

    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    Set CreateOrResetSheet = ws
    Exit Function

EH:
    Application.DisplayAlerts = True
    Err.Raise vbObjectError + 994, , "CreateOrResetSheet failed [" & sheetName & "]: " & Err.Number & " - " & Err.Description
End Function



'------------------------------------------------------------
' Safely sets a PivotTable page filter to a single value.
' - Validates field exists
' - Validates the item exists (prevents silent failure)
' - Clears prior filter properly
'
' @param pt PivotTable to filter
' @param fieldName Pivot page field name (e.g. "Country")
' @param pageValue The PivotItem name to select (e.g. "MY")
'------------------------------------------------------------
Private Sub SafeSetPageFilter(ByVal pt As PivotTable, ByVal fieldName As String, ByVal pageValue As String)
    On Error GoTo EH

    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)

    ' Must be a page field to use CurrentPage reliably
    If pf.Orientation <> xlPageField Then
        Err.Raise vbObjectError + 5001, , _
            "SafeSetPageFilter: [" & fieldName & "] is not a PageField in [" & pt.name & "]"
    End If

    pt.ManualUpdate = True

    ' Clear existing selection
    pf.ClearAllFilters
    pf.EnableMultiplePageItems = False

    ' Validate item exists (prevents "show all" / stale cache behaviour)
    Dim pi As PivotItem
    On Error Resume Next
    Set pi = pf.PivotItems(pageValue)
    On Error GoTo EH

    If pi Is Nothing Then
        Err.Raise vbObjectError + 5002, , _
            "SafeSetPageFilter: value [" & pageValue & "] not found in field [" & fieldName & "] for [" & pt.name & "]"
    End If

    pf.CurrentPage = pageValue

    pt.ManualUpdate = False
    pt.RefreshTable
    DoEvents
    Exit Sub

EH:
    pt.ManualUpdate = False
    Err.Raise vbObjectError + 5003, , "SafeSetPageFilter failed at [" & pt.name & "." & fieldName & "]: " & Err.Number & " - " & Err.Description
End Sub


Private Function SanitizeSheetName(ByVal s As String) As String
    ' Excel sheet name max 31, remove invalid chars: : \ / ? * [ ]
    Dim t As String
    t = s
    t = Replace(t, ":", " ")
    t = Replace(t, "\", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, "?", " ")
    t = Replace(t, "*", " ")
    t = Replace(t, "[", " ")
    t = Replace(t, "]", " ")
    t = Trim$(t)
    If Len(t) > 31 Then t = left$(t, 31)
    SanitizeSheetName = t
End Function


Private Function SanitizeName(ByVal s As String) As String
    ' for pivot table / object names (keep simple)
    Dim t As String
    t = s
    t = Replace(t, " ", "_")
    t = Replace(t, "-", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "\", "_")
    t = Replace(t, "[", "_")
    t = Replace(t, "]", "_")
    t = Replace(t, "(", "_")
    t = Replace(t, ")", "_")
    SanitizeName = t
End Function

'==================================================
' TRUE pivot deletion (not range clear)
'==================================================
Private Sub ClearPivotTablesOnly(ByVal ws As Worksheet)
    On Error Resume Next

    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.ChangePivotCache GetDummyPivotCache()
    Next pt

    Dim i As Long
    For i = ws.PivotTables.Count To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
    Next i

    On Error GoTo 0
End Sub


Private Function RecreateSheet(ByVal sheetName As String, ByVal afterSheet As Worksheet) As Worksheet
    gStep = "RecreateSheet: " & sheetName

    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock("RecreateSheet " & sheetName)

    On Error GoTo EH

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo EH
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
    ws.name = sheetName

    Set RecreateSheet = ws

CleanExit:
    'EndStructureUnlock wasProtected
    Exit Function

EH:
    Application.DisplayAlerts = True
    'EndStructureUnlock wasProtected
    Err.Raise vbObjectError + 993, , "RecreateSheet failed [" & sheetName & "]: " & Err.Number & " - " & Err.Description
End Function



Private Function NextFreePivotRow(ByVal ws As Worksheet, Optional ByVal minRow As Long = 1) As Long
    Dim maxRow As Long
    maxRow = minRow

    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        If Not pt.TableRange2 Is Nothing Then
            maxRow = Application.WorksheetFunction.Max(maxRow, _
                     pt.TableRange2.Row + pt.TableRange2.Rows.Count)
        End If
    Next pt

    NextFreePivotRow = maxRow + 3 ' spacer
End Function

Private Function PivotItemToDate(ByVal pi As PivotItem) As Date
    ' PivotItem.Value often works for date fields; fallback to Name parsing.
    On Error GoTo Fallback

    If IsDate(pi.Value) Then
        PivotItemToDate = CDate(pi.Value)
        Exit Function
    End If

    If IsNumeric(pi.Value) Then
        PivotItemToDate = DateSerial(1899, 12, 30) + CLng(pi.Value)
        Exit Function
    End If

Fallback:
    ' Try Name as date or as serial
    If IsDate(pi.name) Then
        PivotItemToDate = CDate(pi.name)
    ElseIf IsNumeric(pi.name) Then
        PivotItemToDate = DateSerial(1899, 12, 30) + CLng(pi.name)
    Else
        Err.Raise 13, , "Month pivot item not convertible to date: [" & pi.name & "]"
    End If
End Function

'==================================================
' Sheet Ordering
' 1) Dashboard always first
' 2) Country pairs: Dashboard_<country> then Pivots_<country>
' 3) Pivots + Data always last (end)
'==================================================
Private Sub EnforceSheetPositions()
    gStep = "EnforceSheetPositions: start"

    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock("EnforceSheetPositions")

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wsDash As Worksheet, wsData As Worksheet, wsPiv As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsPiv = ThisWorkbook.Worksheets("Pivots")

    ' --- 1) Dashboard first ---
    wsDash.Move Before:=ThisWorkbook.Worksheets(1)

    ' --- 2) Build a sorted list of countries from Dashboard_* ---
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase$(left$(ws.name, 10)) = "dashboard_" Then
            Dim ctry As String
            ctry = Mid$(ws.name, 11) ' after "Dashboard_"
            If Len(ctry) > 0 Then
                If Not dict.Exists(ctry) Then dict.Add ctry, True
            End If
        End If
    Next ws

    Dim countries() As String
    If dict.Count > 0 Then
        ReDim countries(1 To dict.Count)
        Dim i As Long: i = 0
        Dim k As Variant
        For Each k In dict.keys
            i = i + 1
            countries(i) = CStr(k)
        Next k
        QuickSortString countries, LBound(countries), UBound(countries)
    End If

    ' --- 3) Move pairs in order right after Dashboard ---
    Dim anchor As Worksheet
    Set anchor = wsDash

    If dict.Count > 0 Then
        For i = LBound(countries) To UBound(countries)
            Dim dn As String, pn As String
            dn = "Dashboard_" & countries(i)
            pn = "Pivots_" & countries(i)

            On Error Resume Next
            Dim wsd As Worksheet, wsp As Worksheet
            Set wsd = ThisWorkbook.Worksheets(dn)
            Set wsp = ThisWorkbook.Worksheets(pn)
            On Error GoTo EH

            If Not wsd Is Nothing Then
                wsd.Move After:=anchor
                Set anchor = wsd
            End If

            If Not wsp Is Nothing Then
                wsp.Move After:=anchor
                Set anchor = wsp
            End If
        Next i
    End If

    ' --- 4) Pivots + Data last ---
    wsPiv.Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    wsData.Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count - 1)

CleanExit:
    'EndStructureUnlock wasProtected
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    gStep = "EnforceSheetPositions: done"
    Exit Sub

EH:
    'EndStructureUnlock wasProtected
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise vbObjectError + 992, , "EnforceSheetPositions failed: " & Err.Number & " - " & Err.Description
End Sub

'==================================================
' QuickSort for String arrays
'==================================================
Private Sub QuickSortString(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop

        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortString arr, first, j
    If i < last Then QuickSortString arr, i, last
End Sub



Private Sub AssertFieldExists(ByVal pt As PivotTable, ByVal fieldName As String)
    On Error GoTo EH
    Dim tmp As PivotField
    Set tmp = pt.PivotFields(fieldName)
    Exit Sub
EH:
    Err.Raise vbObjectError + 780, , "Pivot field missing: [" & fieldName & "] in PivotTable [" & pt.name & "]"
End Sub
Private Sub ApplyChartValueLabelsAndPercentSeries(ByVal ch As Chart)
    On Error Resume Next

    Dim s As Series
    For Each s In ch.SeriesCollection
        ' Turn on data labels
        s.HasDataLabels = True
        s.DataLabels.ShowValue = True

        Dim nm As String
        nm = LCase$(CStr(s.name))

        Dim isPct As Boolean
        isPct = (InStr(1, nm, "%") > 0) _
                Or (InStr(1, nm, "ach") > 0) _
                Or (InStr(1, nm, "achievement") > 0)

        If isPct Then
            ' Make % series readable: line + secondary axis
            s.AxisGroup = xlSecondary
            s.chartType = xlLineMarkers

            s.DataLabels.NumberFormat = "0.00%"
            s.DataLabels.Position = xlLabelPositionAbove
        Else
            s.DataLabels.NumberFormat = "#,##0"

            ' Better label placement for bars/columns
            If ch.chartType = xlBarClustered Or ch.chartType = xlBarStacked Or ch.chartType = xlBarStacked100 Then
                s.DataLabels.Position = xlLabelPositionOutsideEnd
            Else
                s.DataLabels.Position = xlLabelPositionAbove
            End If
        End If
    Next s

    ' If we created a secondary axis, keep it clean
    If ch.HasAxis(xlValue, xlSecondary) Then
        ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.00%"
    End If

    On Error GoTo 0
End Sub


Private Sub ApplyTableBorders(ByVal rng As Range)
    On Error Resume Next
    If rng Is Nothing Then Exit Sub

    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    rng.Borders(xlEdgeLeft).Weight = xlMedium
    rng.Borders(xlEdgeTop).Weight = xlMedium
    rng.Borders(xlEdgeRight).Weight = xlMedium
    rng.Borders(xlEdgeBottom).Weight = xlMedium
    On Error GoTo 0
End Sub








'-------------------------------------------------------------------------------

 ' Runs a callback Sub while workbook Structure protection is temporarily disabled.
 ' This is the cleanest "wrapper" style.
 '
 ' @param procName The name of a Public Sub with no arguments to call.

'-------------------------------------------------------------------------------

Private Sub RunWithStructureUnlocked(ByVal procName As String)
    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock(procName)

    On Error GoTo EH
    Application.Run procName

CleanExit:
    'EndStructureUnlock wasProtected
    Exit Sub

EH:
    Dim msg As String
    msg = "RunWithStructureUnlocked failed in [" & procName & "]: " & Err.Number & " - " & Err.Description
    'EndStructureUnlock wasProtected
    Err.Raise vbObjectError + 990, , msg
End Sub

'-------------------------------------------------------------------------------
' Ensures tblData has exactly 12 columns (A:L) and the expected headers.
'
' Expected columns:
' 1  Country
' 2  Region
' 3  Center
' 4  Date
' 5  Month
' 6  MonthName
' 7  Budget
' 8  Actual
' 9  Variance
' 10 AchievementPct
' 11 UploadId
' 12 UploadedAt
'-------------------------------------------------------------------------------
Public Sub EnsureTblDataSchema(ByVal wsData As Worksheet, ByVal lo As ListObject)
    gStep = "EnsureTblDataSchema: start"

    ' Hard guarantee (even if caller forgot)
    AssertSheetUnprotected wsData, SYS_SHEET_PWD, "EnsureTblDataSchema"

    lo.ShowHeaders = True

    Do While lo.ListColumns.Count < 12
        lo.ListColumns.Add
    Loop

    Dim hdr As Variant
    hdr = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                "Budget", "Actual", "Variance", "AchievementPct", _
                "UploadId", "UploadedAt")

    lo.HeaderRowRange.Resize(1, 12).Value = hdr

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastRow < lo.Range.Row Then lastRow = lo.Range.Row

    lo.Resize wsData.Range("A1:L" & lastRow)

    ' Formats (safe)
    On Error Resume Next
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "mmm-yyyy"
    lo.ListColumns("Budget").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Actual").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Variance").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("AchievementPct").DataBodyRange.NumberFormat = "0.00%"
    lo.ListColumns("UploadedAt").DataBodyRange.NumberFormat = "yyyy-mm-dd h:mm:ss"
    On Error GoTo 0

    gStep = "EnsureTblDataSchema: done"
End Sub


'-------------------------------------------------------------------------------
' Ensures a worksheet is unprotected (verifies state). Raises if still protected.
'-------------------------------------------------------------------------------
Private Sub AssertSheetUnprotected(ByVal ws As Worksheet, ByVal sheetPwd As String, ByVal stepName As String)

    If DEV_MODE Then Exit Sub

    gStep = stepName & ": Unprotect " & ws.name

    On Error Resume Next
    'ws.Unprotect Password:=sheetPwd
    On Error GoTo 0

    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        Err.Raise vbObjectError + 880, , _
            "Sheet [" & ws.name & "] is still protected after Unprotect()."
    End If
End Sub



Private Sub DeletePivotIfExists(ByVal ws As Worksheet, ByVal ptName As String)
    On Error Resume Next
    ws.PivotTables(ptName).TableRange2.Clear
    On Error GoTo 0
End Sub

'==================================================
' FULL CLEAN: remove ALL pivots from every
' Pivots_<country> worksheet
'
' Use when duplicate pivots keep appearing
'==================================================
Public Sub CleanAllCountryPivotSheets()

    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock("CleanAllCountryPivotSheets")

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

        If LCase$(left$(ws.name, 7)) = "pivots_" Then

            'Debug.Print "Cleaning: " & ws.name

            DetachSlicersFromSheet ws
            HardDeleteAllPivots ws

        End If

    Next ws

    MsgBox "All country pivot sheets cleaned." & vbCrLf & _
           "Now CLOSE EXCEL completely and reopen before importing again.", vbInformation

CleanExit:
    'EndStructureUnlock wasProtected
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    'EndStructureUnlock wasProtected
    MsgBox "Cleanup failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub



Private Sub DetachSlicersFromSheet(ByVal ws As Worksheet)

    Dim sc As SlicerCache
    Dim pt As PivotTable
    Dim i As Long

    For Each sc In ThisWorkbook.SlicerCaches

        For i = sc.PivotTables.Count To 1 Step -1
            Set pt = sc.PivotTables(i)

            If pt.Parent.name = ws.name Then
                sc.PivotTables.RemovePivotTable pt
            End If
        Next i

    Next sc

End Sub
Private Sub HardDeleteAllPivots(ByVal ws As Worksheet)

    On Error Resume Next

    Dim pt As PivotTable
    Dim i As Long

    ' Break cache links first
    For Each pt In ws.PivotTables
        pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=ws.Range("A1:A1"))
    Next pt

    ' Now remove pivots
    For i = ws.PivotTables.Count To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
    Next i

    ws.Cells.ClearFormats

    On Error GoTo 0

End Sub



Private Sub BuildRunBegin()
    gBuildRunId = Format$(Now, "yyyymmdd_hhnnss")
    Set gBuildCountrySeen = CreateObject("Scripting.Dictionary")
    gBuildCountrySeen.CompareMode = vbTextCompare
    'Debug.Print "=== BuildCountryDashboards BEGIN RunId=" & gBuildRunId & " ==="
End Sub

Private Sub BuildRunEnd()
    'Debug.Print "=== BuildCountryDashboards END RunId=" & gBuildRunId & " ==="
    Set gBuildCountrySeen = Nothing
End Sub

Private Function BuildCountryOncePerRun(ByVal country As String) As Boolean
    Dim k As String
    k = NormalizeCountryKey(country)
    If gBuildCountrySeen Is Nothing Then
        BuildCountryOncePerRun = True
        Exit Function
    End If

    If gBuildCountrySeen.Exists(k) Then
        'Debug.Print "SKIP duplicate build (same run): " & k & " RunId=" & gBuildRunId
        BuildCountryOncePerRun = False
    Else
        gBuildCountrySeen.Add k, True
        BuildCountryOncePerRun = True
    End If
End Function

Private Sub DeleteSlicerCacheIfExists(ByVal cacheName As String)
    On Error Resume Next
    ThisWorkbook.SlicerCaches(cacheName).Delete
    On Error GoTo 0
End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub AddCountrySlicers( _
    ByVal wsDash As Worksheet, _
    ByVal ptCenter As PivotTable, _
    ByVal ptCC As PivotTable, _
    ByVal country As String)

    Dim key As String
    key = SanitizeName(country)

    Dim scRegionName As String, slRegionName As String
    Dim scMonthName As String, slMonthName As String

    scRegionName = "scRegion_" & key
    slRegionName = "slRegion_" & key

    scMonthName = "scMonth_" & key
    slMonthName = "slMonth_" & key

    ' Clean old slicers + caches for this country
    DeleteShapeIfExists wsDash, slRegionName
    DeleteShapeIfExists wsDash, slMonthName
    DeleteSlicerCacheIfExists scRegionName
    DeleteSlicerCacheIfExists scMonthName

    ' Create slicer caches from ptCenter (has Region + Month fields)
    Dim scRegion As SlicerCache, scMonth As SlicerCache
    Set scRegion = ThisWorkbook.SlicerCaches.Add2(ptCenter, "Region", scRegionName)
    Set scMonth = ThisWorkbook.SlicerCaches.Add2(ptCenter, "Month", scMonthName)
    ' Ensure slicers hide items with no data (Excel 2019: use numeric enum value)
    On Error Resume Next
    scRegion.CrossFilterType = XL_SLICER_HIDE_BUTTONS_WITH_NO_DATA
    scMonth.CrossFilterType = XL_SLICER_HIDE_BUTTONS_WITH_NO_DATA
    scRegion.ClearManualFilter
    scMonth.ClearManualFilter
    On Error GoTo 0
    
    ' Excel 2019 safe: best-effort hide items without data
    SetSlicerShowItemsWithData scRegion
    SetSlicerShowItemsWithData scMonth
    
    '========================================================
    ' IMPORTANT: Make slicer show ONLY items that have data
    ' under current Country/Month filters (prevents  all regions ).
    '========================================================

    On Error Resume Next
    ptCenter.ManualUpdate = True
    ptCC.ManualUpdate = True
    
    ptCenter.PivotFields("Region").ShowAllItems = False
    ptCenter.PivotFields("Center").ShowAllItems = False
    ptCenter.PivotFields("Month").ShowAllItems = False
    
    ptCC.PivotFields("Region").ShowAllItems = False
    ptCC.PivotFields("Center").ShowAllItems = False
    ptCC.PivotFields("Month").ShowAllItems = False
    
    ptCenter.ManualUpdate = False
    ptCC.ManualUpdate = False
    
    ptCenter.RefreshTable
    ptCC.RefreshTable
    DoEvents
    On Error GoTo 0

    ' Connect slicers to BOTH pivots (chart + breakdown)
    scRegion.PivotTables.AddPivotTable ptCenter
    scRegion.PivotTables.AddPivotTable ptCC

    scMonth.PivotTables.AddPivotTable ptCenter
    scMonth.PivotTables.AddPivotTable ptCC

    ' Add slicers at temp position - RepositionCountryDashboard sets final position
    ' IMPORTANT: rename Shape after Add - Slicer name != Shape name in Excel
    Dim slRegion As Slicer, slMonth As Slicer
    Set slRegion = scRegion.Slicers.Add(wsDash, , slRegionName, "Region", 30, 30, 200, 190)
    Set slMonth = scMonth.Slicers.Add(wsDash, , slMonthName, "Month", 30, 245, 200, 190)
    On Error Resume Next
    slRegion.Shape.name = slRegionName
    slMonth.Shape.name = slMonthName
    On Error GoTo 0

End Sub

'========================================================
' Force combo axis for Country Achievement chart
' Budget/Actual = bars
' Achievement% = line on secondary axis
'========================================================
Public Sub FixCountryAchievementAxis(ByVal ch As Chart)
    On Error GoTo SafeExit

    Dim s As Series
    Dim hasPct As Boolean
    hasPct = False

    For Each s In ch.SeriesCollection
        Dim nm As String
        nm = LCase$(CStr(s.name))

        If InStr(nm, "avg achievement") > 0 Or InStr(nm, "achievement") > 0 Or InStr(nm, "ach %") > 0 Or InStr(nm, "%") > 0 Then
            ' % series as LINE on secondary
            s.AxisGroup = xlSecondary
            s.chartType = xlLineMarkers
            hasPct = True
        Else
            ' budget/actual as BAR on primary
            s.AxisGroup = xlPrimary
            s.chartType = xlBarClustered
        End If
    Next s

    If Not hasPct Then GoTo SafeExit

    ' Force secondary axis to exist
    ch.HasAxis(xlValue, xlSecondary) = True

    ' BAR chart note:
    ' value axis is horizontal; secondary value axis shows at the TOP
    With ch.Axes(xlValue, xlSecondary)
        '.MinimumScale = 0
        '.MaximumScale = 1.2        ' 10%
        '.MajorUnit = 0.1
        .TickLabels.NumberFormat = "0%"

        ' make it visible (TOP)
        .TickLabelPosition = xlTickLabelPositionHigh
        On Error Resume Next
        .Format.Line.Visible = msoTrue
        On Error GoTo 0
    End With

    ' optional: keep primary value axis clean
    On Error Resume Next
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlValue, xlPrimary).Format.Line.Visible = msoTrue
    On Error GoTo 0

    ch.Refresh
SafeExit:
End Sub

'--------------------------------------------------
' Validate tblData headers are exactly what pivots expect.
' Raises a clear error if any required column is missing.
'--------------------------------------------------
Private Sub AssertTblDataHeaders(ByVal lo As ListObject)

    Dim required As Variant
    required = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                     "Budget", "Actual", "Variance", "AchievementPct", _
                     "UploadId", "UploadedAt")

    Dim i As Long
    For i = LBound(required) To UBound(required)
        On Error GoTo Missing
        Dim tmp As ListColumn
        Set tmp = lo.ListColumns(CStr(required(i)))
        
    Next i
    On Error GoTo 0
    Exit Sub

Missing:
    Err.Raise vbObjectError + 881, , _
        "tblData schema mismatch. Missing column header: [" & CStr(required(i)) & "]." & vbCrLf & _
        "Fix: Ensure Data!tblData headers are exactly: " & vbCrLf & _
        Join(required, ", ")
End Sub
'--------------------------------------------------
' Validates that a PivotCache contains all expected fields.
'
' @param pc The PivotCache to validate.
' @param expectedFields A 0-based array of field names (strings).
'--------------------------------------------------


'--------------------------------------------------
' Normalizes tblData header text so PivotCache field names match exactly.
' Fixes: leading/trailing spaces, NBSP (Chr(160)), line breaks.
'--------------------------------------------------
Private Sub NormalizeTblDataHeaders(ByVal lo As ListObject)
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        Dim s As String
        s = CStr(lo.HeaderRowRange.Cells(1, i).Value)

        s = Replace(s, Chr$(160), " ") ' NBSP -> space
        s = Replace(s, vbCr, " ")
        s = Replace(s, vbLf, " ")
        s = Trim$(s)

        lo.HeaderRowRange.Cells(1, i).Value = s
    Next i
End Sub
Private Sub AssertPivotTableFields(ByVal pt As PivotTable, ByVal expectedFields As Variant)
    Dim f As Variant
    For Each f In expectedFields
        AssertFieldExists pt, CStr(f)
    Next f
End Sub

'==================================================
' Materialize PivotCache fields safely using CacheHost (no delete).
'==================================================
'==================================================
' Materialize PivotCache fields safely using zz_CacheHost (VISIBLE + UNPROTECTED).
' Fixes 1004 "Reference isn't valid" during PivotTables.Add.
'==================================================
'==================================================
' PivotTable name exists anywhere in workbook
'==================================================
Private Function PivotTableNameExists(ByVal ptName As String) As Boolean
    Dim ws As Worksheet, pt As PivotTable
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            If LCase$(pt.name) = LCase$(ptName) Then
                PivotTableNameExists = True
                Exit Function
            End If
        Next pt
    Next ws
    On Error GoTo 0
End Function

'==================================================
' Hard delete pivot table everywhere (if it exists)
'==================================================
Private Sub DeletePivotAnywhere(ByVal ptName As String)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.PivotTables(ptName).TableRange2.Clear
    Next ws
    On Error GoTo 0
End Sub
'==================================================
' Create PivotCache using a staged plain range on zz_CacheHost.
' This avoids PivotCaches.Create 1004 caused by ListObject/internal table state.
'==================================================
Private Function CreatePivotCacheFromTable(ByVal lo As ListObject) As PivotCache
    gStep = "CreatePivotCacheFromTable: start"

    If lo Is Nothing Then Err.Raise vbObjectError + 2113, , "ListObject is Nothing."
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 2116, , "tblData has no rows."

    ' Normalize headers first (prevents  field not found  later)
    NormalizeTblDataHeaders lo

    Dim wasStructProtected As Boolean
    wasStructProtected = ThisWorkbook.ProtectStructure

    ' --- MUST unlock structure for Visible/add sheet actions ---
    On Error Resume Next
    If wasStructProtected Then ThisWorkbook.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
    On Error GoTo 0

    Dim wsHost As Worksheet
    gStep = "CreatePivotCacheFromTable: EnsureCacheHostSheetVisible"
    Set wsHost = EnsureCacheHostSheetVisible()


    ' HARD reset host sheet (prevents PivotCaches.Create 1004)
    gStep = "CreatePivotCacheFromTable: clear host"
    wsHost.Cells.Clear
    wsHost.Cells.ClearFormats
    wsHost.Cells.UnMerge



    ' --- Source range from table (exclude totals row if any) ---
    Dim rngSrc As Range
    Set rngSrc = lo.Range
    If lo.ShowTotals Then
        Set rngSrc = rngSrc.Resize(rngSrc.Rows.Count - 1, rngSrc.Columns.Count)
    End If

    ' --- Stage values to host as a plain range ---
    Dim rngDst As Range
    Set rngDst = wsHost.Range("A1").Resize(rngSrc.Rows.Count, 12)

    gStep = "CreatePivotCacheFromTable: stage values"
    rngDst.Value2 = rngSrc.Value2
    Dim hdr As Variant
    hdr = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                "Budget", "Actual", "Variance", "AchievementPct", _
                "UploadId", "UploadedAt")
    
    wsHost.Range("A1").Resize(1, 12).Value = hdr

    gStep = "CreatePivotCacheFromTable: validate staged range"
    ValidatePivotSourceRange rngDst

    ' --- Create PivotCache (STRING is most reliable; Wizard fallback) ---
    Dim pc As PivotCache
    Dim srcLocal As String
    
    gStep = "CreatePivotCacheFromTable: PivotCaches.Create (string)"
    srcLocal = rngDst.Address(ReferenceStyle:=xlR1C1, External:=True)

    ' 1) Create cache first
    On Error GoTo EH_CREATE
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcLocal)
    On Error GoTo 0
    GoTo OK_CREATE
    
EH_CREATE:
    Dim cNo As Long, cMsg As String
    cNo = Err.Number
    cMsg = Err.Description

    Debug.Print "PivotCaches.Create failed:", cNo, cMsg, " -> using Wizard"
    Err.Clear

    On Error GoTo EH_WIZ
    Set pc = CreatePivotCacheFromRange_Wizard(srcLocal, wsHost)
    On Error GoTo 0
    GoTo OK_CREATE

EH_WIZ:
    Dim wNo As Long, wMsg As String
    wNo = Err.Number
    wMsg = Err.Description
    On Error GoTo 0

    Err.Raise vbObjectError + 2117, , _
        "CreatePivotCacheFromTable failed (Create + Wizard)." & vbCrLf & _
        "CreateErr=" & cNo & " - " & cMsg & vbCrLf & _
        "WizardErr=" & wNo & " - " & wMsg
        
OK_CREATE:
    MaterializePivotCacheFields_FromSourceString srcLocal
    
    ' Hide host back (optional but recommended)
    gStep = "CreatePivotCacheFromTable: hide host"
    wsHost.Visible = xlSheetVeryHidden

    ' Restore structure protection if needed
    If wasStructProtected Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=modSecurity.SYS_STRUCT_PWD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If

    gStep = "CreatePivotCacheFromTable: done"
    Set CreatePivotCacheFromTable = pc
End Function ' Validates pivot source range for known PivotCache 1004 triggers
'==================================================
Private Sub ValidatePivotSourceRange(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 2120, , "Pivot source range is Nothing."

    ' 1) Merged cells inside the source range can cause PivotCache 1004
    If rng.MergeCells Then
        Err.Raise vbObjectError + 2121, , "Pivot source contains merged cells. Unmerge Data range."
    End If

    ' 2) Header row cannot contain blanks or errors
    Dim c As Range
    For Each c In rng.Rows(1).Cells
        If IsError(c.Value) Then
            Err.Raise vbObjectError + 2122, , "Pivot header has error at " & c.Address(External:=False)
        End If
        If Trim$(CStr(c.Value)) = "" Then
            Err.Raise vbObjectError + 2123, , "Pivot header is blank at " & c.Address(External:=False)
        End If
    Next c
End Sub

'==================================================
' Ensures a permanent very-hidden sheet exists to host temp pivots/caches.
' This avoids Worksheet.Delete 1004 caused by PivotCache busy state.
'==================================================
Private Function EnsureCacheHostSheet() As Worksheet
    Const SH_NAME As String = "zz_CacheHost"
    Dim ws As Worksheet
    Dim wasProtected As Boolean

    'wasProtected = BeginStructureUnlock("EnsureCacheHostSheet")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_NAME
    End If

    ws.Visible = xlSheetVeryHidden

    'EndStructureUnlock wasProtected
    Set EnsureCacheHostSheet = ws
End Function

Private Function EnsureCacheHostSheetVisible() As Worksheet
    Const SH_NAME As String = "zz_CacheHost"

    Dim wasStructProtected As Boolean
    wasStructProtected = ThisWorkbook.ProtectStructure

    On Error Resume Next
    If wasStructProtected Then ThisWorkbook.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
    On Error GoTo 0

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_NAME
    End If

    ' Must be visible + unprotected for staging/cache creation
    ws.Visible = xlSheetVisible
    modSecurity.UnprotectSheet ws
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        Err.Raise vbObjectError + 9941, , "Sheet still protected: " & ws.name
    End If

    If wasStructProtected Then
        ThisWorkbook.Protect Password:=modSecurity.SYS_STRUCT_PWD, Structure:=True, Windows:=False
    End If

    Set EnsureCacheHostSheetVisible = ws
End Function

Private Function BuildCountryCacheHostSheetName(ByVal countryKey As String) As String
    Dim nm As String
    nm = "zz_CH_" & SanitizeSheetName(NormalizeCountryKey(countryKey))
    If Len(nm) = 0 Then nm = "zz_CH_UNK"
    If Len(nm) > 31 Then nm = left$(nm, 31)
    BuildCountryCacheHostSheetName = nm
End Function

Private Function EnsureCountryCacheHostSheetVisible(ByVal countryKey As String) As Worksheet
    Dim shName As String
    shName = BuildCountryCacheHostSheetName(countryKey)

    Dim wasStructProtected As Boolean
    wasStructProtected = ThisWorkbook.ProtectStructure

    On Error Resume Next
    If wasStructProtected Then ThisWorkbook.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
    On Error GoTo 0

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = shName
    End If

    ws.Visible = xlSheetVisible
    modSecurity.UnprotectSheet ws
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        Err.Raise vbObjectError + 9942, , "Sheet still protected: " & ws.name
    End If

    If wasStructProtected Then
        ThisWorkbook.Protect Password:=modSecurity.SYS_STRUCT_PWD, Structure:=True, Windows:=False
    End If

    Set EnsureCountryCacheHostSheetVisible = ws
End Function

Private Function BuildPivotSourceR1C1(ByVal rng As Range) As String
    BuildPivotSourceR1C1 = "'" & rng.Parent.name & "'!" & _
        rng.Address(ReferenceStyle:=xlR1C1, External:=False)
End Function

Private Function GetDummyPivotCache() As PivotCache
    
    ' Create ONCE and reuse forever (prevents PivotCaches.Count explosion)
    If Not gDummyPivotCache Is Nothing Then
        Set GetDummyPivotCache = gDummyPivotCache
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = EnsureCacheHostSheetVisible() ' you already have this helper

    ws.Cells.Clear
    ws.Range("A1").Value = "X"

    Set gDummyPivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1:A1") _
    )

    Set GetDummyPivotCache = gDummyPivotCache
End Function

'==================================================
' Ensure a worksheet exists (create if missing)
'==================================================
Private Function EnsureSheet(ByVal sheetName As String, ByVal afterSheet As Worksheet) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Dim wasProtected As Boolean
        'wasProtected = BeginStructureUnlock("EnsureSheet " & sheetName)

        Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        ws.name = sheetName

        'EndStructureUnlock wasProtected
    End If

    Set EnsureSheet = ws
End Function

'==================================================
' Ensure Calc exists + set it to very hidden
'==================================================
Private Function EnsureCalcSheet() As Worksheet
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    Dim ws As Worksheet
    Set ws = EnsureSheet("Calc", wsDash)

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0

    Set EnsureCalcSheet = ws
End Function

Private Function GetPivotOrFail(ByVal ws As Worksheet, ByVal ptName As String, ByVal stepName As String) As PivotTable
    On Error GoTo EH
    Dim pt As PivotTable
    Set pt = ws.PivotTables(ptName)
    Set GetPivotOrFail = pt
    Exit Function
EH:
    Dim names As String, p As PivotTable
    names = ""
    For Each p In ws.PivotTables
        names = names & p.name & ", "
    Next p
    If Len(names) > 2 Then names = left$(names, Len(names) - 2)

    Err.Raise vbObjectError + 8801, , _
        stepName & ": PivotTable not found [" & ptName & "] on sheet [" & ws.name & "]." & vbCrLf & _
        "Existing pivots: " & IIf(Len(names) = 0, "(none)", names)
End Function
'==================================================
' Safe wrapper: only fix axis if the chart exists and is a real chart
'==================================================
Private Sub TryFixCountryAchievementAxis(ByVal wsDash As Worksheet, ByVal chartShapeName As String)
    On Error GoTo SafeExit
    
    Dim shp As Shape
    Set shp = wsDash.Shapes(chartShapeName)  ' <-- will throw if not found
    
    ' If it's not a chart, do nothing
    If shp.HasChart Then
        FixCountryAchievementAxis shp.Chart
    End If

SafeExit:
End Sub
Private Sub AssertSamePivotCache(ByVal basePt As PivotTable, ByVal otherPt As PivotTable, ByVal stepName As String)
    If basePt.PivotCache.Index <> otherPt.PivotCache.Index Then
        Err.Raise vbObjectError + 7771, , stepName & _
            ": PivotCache mismatch. " & basePt.name & "(Cache=" & basePt.PivotCache.Index & _
            ") vs " & otherPt.name & "(Cache=" & otherPt.PivotCache.Index & ")"
    End If
End Sub

 Private Sub ForceAchievementSecondaryAxis(ByVal ch As Chart)
     On Error GoTo EH
 
     ' Step 1: Reset entire chart to Column so Excel allows combo
     ch.chartType = xlColumnClustered
     DoEvents
 
     Dim s As Series
     Dim hasSecondary As Boolean
     hasSecondary = False
 
     ' Step 2: First pass - assign axis groups BEFORE changing per-series chartType
     For Each s In ch.SeriesCollection
         Dim nm As String
         nm = LCase$(CStr(s.name))
         If InStr(nm, "ach") > 0 Or InStr(nm, "%") > 0 Or InStr(nm, "achievement") > 0 Then
             s.AxisGroup = xlSecondary
             hasSecondary = True
         Else
             s.AxisGroup = xlPrimary
         End If
     Next s
 
     ' Step 3: Force secondary axis to exist BEFORE changing series chartType
     If hasSecondary Then
         ch.HasAxis(xlValue, xlSecondary) = True
         DoEvents
     End If
 
     ' Step 4: Second pass - change chartType per series (axis must exist first)
     For Each s In ch.SeriesCollection
         nm = LCase$(CStr(s.name))
         If InStr(nm, "ach") > 0 Or InStr(nm, "%") > 0 Or InStr(nm, "achievement") > 0 Then
             s.chartType = xlLineMarkers
         Else
             s.chartType = xlColumnClustered
         End If
     Next s
 
     ' Step 5: Format secondary axis
     If hasSecondary Then
         With ch.Axes(xlValue, xlSecondary)
             .TickLabels.NumberFormat = "0%"
             .HasTitle = False
         End With
     End If
 
     ch.Refresh
 
     ' Step 6: Re-apply number format after Refresh (pivot chart resets label formats)
     Dim sPost As Series
     For Each sPost In ch.SeriesCollection
         Dim nmPost As String
         nmPost = LCase$(CStr(sPost.name))
         If InStr(nmPost, "ach") = 0 And InStr(nmPost, "%") = 0 And InStr(nmPost, "achievement") = 0 Then
             If sPost.HasDataLabels Then
                 sPost.DataLabels.NumberFormat = "#,##0"
             End If
         End If
     Next sPost
     Exit Sub
 
EH:
     ' swallow (chart can be finicky during rebuild)
 End Sub

Public Sub FixCountryAchievementAxis_CustomCombo(ByVal ch As Chart)
    On Error GoTo SafeExit

    Dim s As Series
    Dim pctSeriesFound As Boolean
    pctSeriesFound = False

    ' IMPORTANT: base must be Column to match Excel Combo UI expectation
    ch.chartType = xlColumnClustered

    For Each s In ch.SeriesCollection
        Dim nm As String
        nm = LCase$(CStr(s.name))

        If InStr(nm, "avg achievement") > 0 Or InStr(nm, "achievement") > 0 Or InStr(nm, "ach") > 0 Or InStr(nm, "%") > 0 Then
            ' Avg Achievement % -> Line on Secondary axis
            s.chartType = xlLine
            s.AxisGroup = xlSecondary
            pctSeriesFound = True

            s.HasDataLabels = True
            s.DataLabels.NumberFormat = "0.00%"
            On Error Resume Next
            s.DataLabels.Position = xlLabelPositionAbove
            On Error GoTo 0
        Else
            ' Budget/Actual -> Clustered Column on Primary
            s.chartType = xlColumnClustered
            s.AxisGroup = xlPrimary

            s.HasDataLabels = True
            s.DataLabels.NumberFormat = "#,##0"
            On Error Resume Next
            s.DataLabels.Position = xlLabelPositionOutsideEnd
            On Error GoTo 0
        End If
    Next s

    If pctSeriesFound Then
        ' Force secondary axis to exist (right side)
        ch.HasAxis(xlValue, xlSecondary) = True

        With ch.Axes(xlValue, xlSecondary)
            '.MinimumScale = 0
            '.MaximumScale = 1          ' 100%
            '.MajorUnit = 0.1
            .TickLabels.NumberFormat = "0.00%"
        End With
    End If

    ' Optional: make columns look nicer (like Excel default combo)
    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 120
    On Error GoTo 0

    ch.Refresh

SafeExit:
End Sub

Public Sub BeginUnlockedBuild()
    If modSecurity.DEV_MODE Then Exit Sub

    ' Allow sheet delete/rename/recreate ONLY during build
    modSecurity.UnprotectWorkbookStructure

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Unprotect everything during rebuild (safe)
        modSecurity.UnprotectSheet ws
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            Err.Raise vbObjectError + 9941, , "Sheet still protected: " & ws.name
        End If
    Next ws
    
    On Error GoTo 0
End Sub

Public Sub EndUnlockedBuild()
    If modSecurity.DEV_MODE Then Exit Sub

    ' Re-protect everything (safe even if sheets were deleted/recreated)
    On Error Resume Next
    modSecurity.ApplyHardening
    On Error GoTo 0
End Sub
'--------------------------------------------------
' Builds a unique upload id for the current import run.
'
' @param uploadedAt The timestamp of the import run.
' @return Unique UploadId string
'--------------------------------------------------
Public Function BuildUploadId(ByVal uploadedAt As Date) As String
    ' Example: UPL_20260222_184501_123_0456
    Randomize
    BuildUploadId = "UPL_" & Format$(uploadedAt, "yyyymmdd_hhnnss") & _
                    "_" & Format$(CLng((Timer * 1000) Mod 1000), "000") & _
                    "_" & Format$(CLng(Int((9999 - 1000 + 1) * Rnd + 1000)), "0000")
End Function
'--------------------------------------------------
' Builds/refreshes Upload Summary table.
' Output columns:
' Country, Region, Center, Month, LatestDay, LatestUploadedAt
'--------------------------------------------------
Public Sub BuildUploadSummary()
    gStep = "BuildUploadSummary: start"
    On Error GoTo EH

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        gStep = "BuildUploadSummary: tblData empty, ensure sheet/table then exit"
        EnsureUploadSummarySheetAndTable
        ClearUploadSummaryBody
        Exit Sub
    End If

    Dim wsOut As Worksheet
    Dim loOut As ListObject
    Set loOut = EnsureUploadSummarySheetAndTable()

    ' Key: Country|Region|Center|MonthSerial
    ' Value: Array(latestDay As Date, latestUploadedAt As Date)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    Dim nRows As Long
    nRows = lo.DataBodyRange.Rows.Count

    For r = 1 To nRows
        Dim country As String, region As String, center As String
        Dim vDate As Variant, vMonth As Variant, vUpAt As Variant

        country = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value)) ' Country
        region = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value))  ' Region
        center = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value))  ' Center
        vDate = lo.DataBodyRange.Cells(r, 4).Value                ' Date
        vMonth = lo.DataBodyRange.Cells(r, 5).Value               ' Month (DateSerial)
        vUpAt = lo.DataBodyRange.Cells(r, 12).Value               ' UploadedAt

        If Len(country) = 0 Or Len(region) = 0 Or Len(center) = 0 Then GoTo NextR
        If Not IsDate(vDate) Then GoTo NextR
        If Not IsDate(vMonth) Then GoTo NextR
        If Not IsDate(vUpAt) Then GoTo NextR

        Dim d As Date, m As Date, upAt As Date
        d = CDate(vDate)
        m = CDate(vMonth)
        upAt = CDate(vUpAt)

        Dim k As String
        k = country & "|" & region & "|" & center & "|" & Format$(m, "yyyymmdd")

        If Not dict.Exists(k) Then
            dict.Add k, Array(d, upAt)
        Else
            Dim arr As Variant
            arr = dict(k)

            If d > arr(0) Then arr(0) = d
            If upAt > arr(1) Then arr(1) = upAt

            dict(k) = arr
        End If

NextR:
    Next r

    ' Clear output and write fresh
    ClearUploadSummaryBody

    If dict.Count = 0 Then
        gStep = "BuildUploadSummary: no valid rows, done"
        Exit Sub
    End If

    ' Write to table body using array for speed
    Dim outArr() As Variant
    ReDim outArr(1 To dict.Count, 1 To 6)

    Dim i As Long: i = 0
    Dim key As Variant
    For Each key In dict.keys
        i = i + 1

        Dim parts() As String
        parts = Split(CStr(key), "|")

        Dim vals As Variant
        vals = dict(key)

        outArr(i, 1) = parts(0)                 ' Country
        outArr(i, 2) = parts(1)                 ' Region
        outArr(i, 3) = parts(2)                 ' Center
        outArr(i, 4) = DateSerial(CInt(left$(parts(3), 4)), CInt(Mid$(parts(3), 5, 2)), CInt(Right$(parts(3), 2))) ' Month
        outArr(i, 5) = vals(0)                  ' LatestDay
        outArr(i, 6) = vals(1)                  ' LatestUploadedAt
    Next key

    ' Ensure table has enough rows then write
    Dim startCell As Range
    Set startCell = loOut.HeaderRowRange.Cells(1, 1).Offset(1, 0)

    startCell.Resize(UBound(outArr, 1), UBound(outArr, 2)).Value = outArr

    ' Resize table to include new body
    Dim lastRow As Long
    lastRow = loOut.HeaderRowRange.Row + UBound(outArr, 1)
    loOut.Resize loOut.Parent.Range(loOut.HeaderRowRange.Cells(1, 1), loOut.Parent.Cells(lastRow, 6))

    ' Formats
    On Error Resume Next
    loOut.ListColumns("Month").DataBodyRange.NumberFormat = "mmm-yyyy"
    loOut.ListColumns("LatestDay").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    loOut.ListColumns("LatestUploadedAt").DataBodyRange.NumberFormat = "yyyy-mm-dd hh:nn:ss"
    On Error GoTo 0

    ' Sort (Country, Region, Center, Month)
    SortUploadSummary loOut

    gStep = "BuildUploadSummary: done"
    Exit Sub

EH:
    Err.Raise vbObjectError + 7601, , gStep & ": " & Err.Number & " - " & Err.Description
End Sub

'--------------------------------------------------
' Builds/refreshes a pivot sheet that shows the latest
' transaction Date by Country for the latest Month.
'
' Layout target:
' A1: "Latest update as at yyyy-mm-dd hh:nn:ss"
' A3:B3: Month report filter
' A5: Pivot body (Country | Max of Date)
'--------------------------------------------------
Private Sub BuildLatestUpdatePivot()
    gStep = "BuildLatestUpdatePivot: start"
    On Error GoTo EH

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")

    Dim wsOut As Worksheet
    Set wsOut = EnsureLatestUpdateSheet()
    wsOut.Visible = xlSheetVisible
    wsOut.Cells.Clear

    Dim asAt As Date
    asAt = GetLatestUploadedAtFromTblData(lo)
    If asAt = 0 Then asAt = Now

    With wsOut.Range("A1")
        .Value = "Latest update as at " & Format$(asAt, "yyyy-mm-dd hh:nn:ss")
        .Font.Bold = True
        .Font.size = 14
    End With

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        wsOut.Range("A5").Value = "No data found in tblData."
        wsOut.Columns("A:B").AutoFit
        Exit Sub
    End If

    ' Ensure pivot name is unique before rebuild.
    DeletePivotAnywhere "ptLatestUpdate"

    Dim pc As PivotCache
    Set pc = CreatePivotCacheFromTable(lo)

    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsOut.Range("A5"), TableName:="ptLatestUpdate")

    With pt
        .ManualUpdate = True
        .ClearAllFilters

        With .PivotFields("Month")
            .Orientation = xlPageField
            .Position = 1
            .NumberFormat = "mmm-yyyy"
        End With

        With .PivotFields("Country")
            .Orientation = xlRowField
            .Position = 1
        End With

        .AddDataField .PivotFields("Date"), "Max of Date", xlMax
        .DataFields("Max of Date").NumberFormat = "yyyy-mm-dd"
        .RowAxisLayout xlTabularRow

        On Error Resume Next
        .TableStyle2 = "PivotStyleLight16"
        On Error GoTo EH

        .ManualUpdate = False
    End With

    pt.RefreshTable
    DoEvents

    Dim latestMonth As Date
    latestMonth = GetLatestMonthFromTblData(lo)
    If latestMonth > 0 Then
        ApplyLatestMonthFilter pt, latestMonth
    End If

    On Error Resume Next
    wsOut.Columns("A:B").AutoFit
    On Error GoTo 0

    gStep = "BuildLatestUpdatePivot: done"
    Exit Sub

EH:
    Err.Raise vbObjectError + 7610, , gStep & ": " & Err.Number & " - " & Err.Description
End Sub

'--------------------------------------------------
' Ensures "Latest Update" sheet exists.
'--------------------------------------------------
Private Function EnsureLatestUpdateSheet() As Worksheet
    On Error GoTo EH
    Const WS_NAME As String = "Latest Update"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_NAME)
    On Error GoTo EH

    If ws Is Nothing Then
        Dim anchor As Worksheet
        On Error Resume Next
        Set anchor = ThisWorkbook.Worksheets("Upload Summary")
        On Error GoTo EH
        If anchor Is Nothing Then Set anchor = ThisWorkbook.Worksheets("Dashboard")

        Set ws = ThisWorkbook.Worksheets.Add(After:=anchor)
        ws.name = WS_NAME
    End If

    Set EnsureLatestUpdateSheet = ws
    Exit Function

EH:
    Err.Raise vbObjectError + 7611, , "EnsureLatestUpdateSheet: " & Err.Number & " - " & Err.Description
End Function

'--------------------------------------------------
' Returns max UploadedAt from tblData (col 12), 0 if none.
'--------------------------------------------------
Private Function GetLatestUploadedAtFromTblData(ByVal lo As ListObject) As Date
    On Error GoTo SafeExit
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim best As Date
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim v As Variant
        v = lo.DataBodyRange.Cells(r, 12).Value
        If IsDate(v) Then
            If CDate(v) > best Then best = CDate(v)
        End If
    Next r
    GetLatestUploadedAtFromTblData = best
SafeExit:
End Function

'--------------------------------------------------
' Returns latest Month (normalized to month start), 0 if none.
'--------------------------------------------------
Private Function GetLatestMonthFromTblData(ByVal lo As ListObject) As Date
    On Error GoTo SafeExit
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Long
    Dim best As Date
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim v As Variant
        v = lo.DataBodyRange.Cells(r, 5).Value
        If IsDate(v) Then
            Dim m As Date
            m = DateSerial(Year(CDate(v)), Month(CDate(v)), 1)
            If m > best Then best = m
        End If
    Next r
    GetLatestMonthFromTblData = best
SafeExit:
End Function

'--------------------------------------------------
' Best-effort set of the Month report filter to latest month.
'--------------------------------------------------
Private Sub ApplyLatestMonthFilter(ByVal pt As PivotTable, ByVal latestMonth As Date)
    On Error GoTo SafeExit
    If pt Is Nothing Then Exit Sub

    Dim pf As PivotField
    Set pf = pt.PivotFields("Month")
    If pf Is Nothing Then Exit Sub

    Dim tgt As Date
    tgt = DateSerial(Year(latestMonth), Month(latestMonth), 1)

    pf.ClearAllFilters
    pf.EnableMultiplePageItems = False

    On Error Resume Next
    pf.CurrentPage = Format$(tgt, "mmm-yyyy")
    If Err.Number = 0 Then
        On Error GoTo SafeExit
        Exit Sub
    End If
    Err.Clear
    On Error GoTo SafeExit

    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        If LCase$(Trim$(pi.name)) <> "(blank)" Then
            Dim matched As Boolean
            matched = False

            On Error Resume Next
            If IsDate(pi.Value) Then
                matched = (DateSerial(Year(CDate(pi.Value)), Month(CDate(pi.Value)), 1) = tgt)
            ElseIf IsDate(pi.name) Then
                matched = (DateSerial(Year(CDate(pi.name)), Month(CDate(pi.name)), 1) = tgt)
            End If
            On Error GoTo SafeExit

            If matched Then
                On Error Resume Next
                pf.CurrentPage = pi.name
                Err.Clear
                On Error GoTo SafeExit
                Exit For
            End If
        End If
    Next pi

SafeExit:
End Sub


'--------------------------------------------------
' Ensures "Upload Summary" sheet + "tblUploadSummary" exist.
'
' @return The ensured output ListObject
'--------------------------------------------------
Private Function EnsureUploadSummarySheetAndTable() As ListObject
    gStep = "EnsureUploadSummarySheetAndTable: start"
    On Error GoTo EH

    Const WS_NAME As String = "Upload Summary"
    Const LO_NAME As String = "tblUploadSummary"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_NAME)
    On Error GoTo EH

    If ws Is Nothing Then
        Dim anchor As Worksheet
        Set anchor = ThisWorkbook.Worksheets("Dashboard")

        Set ws = ThisWorkbook.Worksheets.Add(After:=anchor)
        ws.name = WS_NAME
    End If

    ws.Visible = xlSheetVisible

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(LO_NAME)
    On Error GoTo EH

    If lo Is Nothing Then
        ' Build headers at A1:F1
        ws.Range("A1:F1").Value = Array("Country", "Region", "Center", "Month", "LatestDay", "LatestUploadedAt")

        Set lo = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:F2"), _
            XlListObjectHasHeaders:=xlYes)

        lo.name = LO_NAME
    Else
        ' Ensure headers correct
        lo.HeaderRowRange.Resize(1, 6).Value = Array("Country", "Region", "Center", "Month", "LatestDay", "LatestUploadedAt")
    End If

    ' Cosmetic
    On Error Resume Next
    ws.Columns("A:F").AutoFit
    On Error GoTo 0

    Set EnsureUploadSummarySheetAndTable = lo
    gStep = "EnsureUploadSummarySheetAndTable: done"
    Exit Function

EH:
    Err.Raise vbObjectError + 7602, , gStep & ": " & Err.Number & " - " & Err.Description
End Function


'--------------------------------------------------
' Clears the body rows of tblUploadSummary safely.
'--------------------------------------------------
Private Sub ClearUploadSummaryBody()
    gStep = "ClearUploadSummaryBody: start"
    On Error GoTo SafeExit

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Upload Summary")

    Dim lo As ListObject
    Set lo = ws.ListObjects("tblUploadSummary")

    If lo Is Nothing Then GoTo SafeExit

    Application.DisplayAlerts = False

    ' Keep table stable: keep exactly 1 blank data row (no DataBodyRange.Delete)
    Dim i As Long
    If lo.ListRows.Count = 0 Then
        lo.ListRows.Add
    Else
        For i = lo.ListRows.Count To 2 Step -1
            lo.ListRows(i).Delete
        Next i
    End If

    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearContents
    End If

SafeExit:
    Application.DisplayAlerts = True
End Sub

'--------------------------------------------------
' Sorts tblUploadSummary by Country, Region, Center, Month.
'--------------------------------------------------
Private Sub SortUploadSummary(ByVal lo As ListObject)
    On Error GoTo SafeExit
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=lo.ListColumns("Country").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=lo.ListColumns("Region").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=lo.ListColumns("Center").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=lo.ListColumns("Month").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

SafeExit:
End Sub

'========================================================
' Deletes all shapes on a sheet with the given name.
' Works even if Excel ended up with duplicated names.
'
' @param ws Target worksheet
' @param shapeName Exact shape name to delete
' @return Number of shapes deleted
'========================================================
Private Function DeleteAllShapesByName(ByVal ws As Worksheet, ByVal shapeName As String) As Long
    Dim i As Long, n As Long

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If LCase$(ws.Shapes(i).name) = LCase$(shapeName) Then
            ws.Shapes(i).Delete
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
    Next i
    On Error GoTo 0

    DeleteAllShapesByName = n
End Function


'------------------------------------------------------------
' Wait until Excel considers the workbook "Saved" (no pending changes).
'
' @param wb Workbook to check.
' @param timeoutSec Max seconds to wait.
' @return True if saved, False if timeout.
'------------------------------------------------------------
Public Function WaitForWorkbookSaved(ByVal wb As Workbook, ByVal timeoutSec As Long) As Boolean
    Dim t0 As Single: t0 = Timer

    Do
        DoEvents
        If wb.Saved Then
            WaitForWorkbookSaved = True
            Exit Function
        End If
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop While (Timer - t0) < timeoutSec

    WaitForWorkbookSaved = False
End Function

'------------------------------------------------------------
' Wait until the file's size & modified time remain unchanged for stableSec.
' This is a best-effort proxy that the file is fully written to disk.
'
' @param filePath Full path to file.
' @param stableSec Required consecutive seconds stable.
' @param timeoutSec Max seconds to wait.
' @return True if stable, False if timeout or file missing.
'------------------------------------------------------------
Public Function WaitForFileStable(ByVal filePath As String, ByVal stableSec As Long, ByVal timeoutSec As Long) As Boolean
    On Error GoTo EH

'    If Len(Dir(filePath)) = 0 Then
'        WaitForFileStable = False
'        Exit Function
'    End If

    If Not FileExistsSmart(filePath) Then
        Err.Raise vbObjectError + 9001, , "File does not exist or not accessible:" & vbCrLf & filePath
    End If

    Dim t0 As Single: t0 = Timer
    Dim stableCount As Long: stableCount = 0

    Dim lastSize As Currency
    Dim lastTime As Date
    GetFileState filePath, lastSize, lastTime

    Do
        DoEvents

        Dim curSize As Currency
        Dim curTime As Date
        GetFileState filePath, curSize, curTime

        If (curSize = lastSize) And (curTime = lastTime) Then
            stableCount = stableCount + 1
            If stableCount >= stableSec Then
                WaitForFileStable = True
                Exit Function
            End If
        Else
            stableCount = 0
            lastSize = curSize
            lastTime = curTime
        End If

        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop While (Timer - t0) < timeoutSec

    WaitForFileStable = False
    Exit Function

EH:
    WaitForFileStable = False
End Function

'------------------------------------------------------------
' Get file size + last modified.
' Uses FileLen + FileDateTime (simple & reliable for local synced folders).
'------------------------------------------------------------
Private Sub GetFileState(ByVal filePath As String, ByRef sizeOut As Currency, ByRef timeOut As Date)
    sizeOut = FileLen(filePath)
    timeOut = FileDateTime(filePath)
End Sub

Public Function FileExistsSmart(ByVal filePath As String) As Boolean
    On Error GoTo Fail

    '------------------------------------------------
    ' Case 1 : HTTP / SharePoint / OneDrive URL
    '------------------------------------------------
    If LCase$(left$(filePath, 4)) = "http" Then
        
        Dim wb As Workbook
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        
        ' try silent open
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
        
        If Not wb Is Nothing Then
            wb.Close False
            FileExistsSmart = True
            Exit Function
        End If
    End If

    '------------------------------------------------
    ' Case 2 : Local filesystem
    '------------------------------------------------
    If Len(Dir$(filePath, vbNormal)) > 0 Then
        FileExistsSmart = True
        Exit Function
    End If

Fail:
    FileExistsSmart = False
End Function

'------------------------------------------------------------
' Dynamically scale % secondary axis based on actual data
'------------------------------------------------------------
Public Sub AutoScalePercentAxis(ByVal ch As Chart)

    On Error GoTo SafeExit

    Dim s As Series
    Dim maxVal As Double
    maxVal = 0

    ' find highest % across all % series
    For Each s In ch.SeriesCollection
        If s.AxisGroup = xlSecondary Then

            Dim v As Variant
            v = s.Values

            Dim i As Long
            For i = 1 To UBound(v)
                If IsNumeric(v(i)) Then
                    If v(i) > maxVal Then maxVal = v(i)
                End If
            Next i

        End If
    Next s

    If maxVal = 0 Then Exit Sub

    ' add headroom (15%)
    maxVal = maxVal * 1.15

    ' round nicely (0.05 = 5% steps)
    Dim roundedMax As Double
    roundedMax = WorksheetFunction.Ceiling(maxVal, 0.05)

    If roundedMax < 0.1 Then roundedMax = 0.1

    With ch.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = roundedMax
        .MajorUnit = roundedMax / 5
        .TickLabels.NumberFormat = "0%"
    End With

SafeExit:
End Sub
'========================================================
' Format ONLY the Achievement/% series labels:
' - Grey background
' - White text
' - 1 decimal percent (0.0%)
' Does NOT change chart type, axis group, or source data.
'========================================================
Public Sub FormatAchievementLabelsSafe(ByVal ch As Chart)
    On Error GoTo SafeExit
    If ch Is Nothing Then Exit Sub
    If ch.SeriesCollection.Count = 0 Then Exit Sub

    Dim s As Series
    For Each s In ch.SeriesCollection
        If IsAchievementSeries(s) Then
            On Error Resume Next
            s.HasDataLabels = True
            s.DataLabels.ShowValue = True
            s.DataLabels.NumberFormat = "0.0%"   ' <-- 1 decimal percent
            On Error GoTo 0

            ' Format each label (some Excel builds are picky; keep per-label error handling)
            Dim dl As DataLabel
            For Each dl In s.DataLabels
                On Error Resume Next

                ' Grey background
                With dl.Format.Fill
                    .Visible = msoTrue
                    .Solid
                    .ForeColor.RGB = RGB(120, 120, 120)
                    .Transparency = 0#
                End With

                ' White text
                dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

                ' Optional readability tweaks
                dl.Format.TextFrame2.TextRange.Font.size = 9
                dl.Format.Line.Visible = msoFalse

                On Error GoTo 0
            Next dl
        End If
    Next s

SafeExit:
End Sub

'========================================================
' Detect the Achievement/% series robustly.
' Prefer: secondary axis OR name contains achievement hints.
'========================================================
Private Function IsAchievementSeries(ByVal s As Series) As Boolean
    On Error GoTo Fail

    If s.AxisGroup = xlSecondary Then
        IsAchievementSeries = True
        Exit Function
    End If

    Dim nm As String
    nm = LCase$(CStr(s.name))

    IsAchievementSeries = (InStr(1, nm, "ach", vbTextCompare) > 0) _
                          Or (InStr(1, nm, "achievement", vbTextCompare) > 0) _
                          Or (InStr(1, nm, "%", vbTextCompare) > 0)
    Exit Function

Fail:
    IsAchievementSeries = False
End Function

Private Sub RestrictRegionDropdown(ByVal pt As PivotTable)
    On Error Resume Next

    ' 1) CRITICAL: drop stale items from cache (removes AH/HN/SD from the dropdown list after refresh)
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone

    pt.ManualUpdate = True

    ' 2) Ensure Region/Center don't keep "all items"
    With pt.PivotFields("Region")
        .ClearAllFilters
        .ShowAllItems = False
    End With

    With pt.PivotFields("Center")
        .ClearAllFilters
        .ShowAllItems = False
    End With

    pt.ManualUpdate = False

    ' 3) Refresh twice (Excel sometimes needs 2 refreshes to purge cache items)
    pt.RefreshTable
    DoEvents
    pt.RefreshTable
    DoEvents

    On Error GoTo 0
End Sub
Private Function CreatePivotCacheFromRange_Wizard(ByVal srcLocal As String, ByVal wsHost As Worksheet) As PivotCache
    Dim ptTmp As PivotTable
    Dim dst As Range

    ' Put temp pivot far away so it won't clash with your staged data
    Set dst = wsHost.Range("N1")

    ' Clear any old temp pivot footprint
    On Error Resume Next
    wsHost.Range("N1:Z50").Clear
    On Error GoTo 0

    ' PivotTableWizard creates the PivotCache internally (more robust than PivotCaches.Create)
    
    Dim tmpName As String
    tmpName = "pt__tmp_" & Format$(Now, "hhnnss") & "_" & CLng(Timer * 1000)
    
    Set ptTmp = wsHost.PivotTableWizard( _
        SourceType:=xlDatabase, _
        SourceData:=srcLocal, _
        TableDestination:=dst, _
        TableName:=tmpName)
        
    Set CreatePivotCacheFromRange_Wizard = ptTmp.PivotCache

    ' Optional cleanup: remove the temp pivot output (cache remains)
    On Error Resume Next
    ptTmp.TableRange2.Clear
    On Error GoTo 0
End Function
'==================================================
' Materialize PivotCache fields using the SAME clean source string
' that created the cache (zz_CacheHost staged range).
' This avoids pc.SourceData weirdness.
'==================================================
Private Sub MaterializePivotCacheFields_FromSourceString(ByVal srcLocal As String)
    gStep = "MaterializePivotCacheFields_FromSourceString: start"

    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock("MaterializePivotCacheFields_FromSourceString")

    Dim wsHost As Worksheet
    Dim ptTmp As PivotTable
    Dim dst As Range

    On Error GoTo EH

    Set wsHost = EnsureCacheHostSheetVisible()

    ' IMPORTANT: do NOT clear A:L because that is your staged source data
    ' Only clear the temp pivot area
    ClearPivotTablesOnly wsHost
    wsHost.Range("N1:Z200").Clear
    wsHost.Cells.UnMerge

    Set dst = wsHost.Range("N1")

   
    Dim tmpName As String
    tmpName = "pt__tmp_" & Format$(Now, "hhnnss") & "_" & CLng(Timer * 1000)
    
    Set ptTmp = wsHost.PivotTableWizard( _
        SourceType:=xlDatabase, _
        SourceData:=srcLocal, _
        TableDestination:=dst, _
        TableName:=tmpName)
    

    ptTmp.ManualUpdate = False
    ptTmp.RefreshTable
    DoEvents

    ' Optional: clear pivot footprint only (leave staged data)
    On Error Resume Next
    ptTmp.TableRange2.Clear
    On Error GoTo 0

CleanExit:
    On Error Resume Next
    wsHost.Visible = xlSheetVeryHidden
    On Error GoTo 0

    'EndStructureUnlock wasProtected
    gStep = "MaterializePivotCacheFields_FromSourceString: done"
    Exit Sub

EH:
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description

    On Error Resume Next
    wsHost.Visible = xlSheetVisible
    On Error GoTo 0

    'EndStructureUnlock wasProtected
    Err.Raise vbObjectError + 2102, , "MaterializePivotCacheFields_FromSourceString failed: " & errNo & " - " & errMsg
End Sub
'========================================================
' RUN-LEVEL UNLOCK/LOCK (NO mid-run protection changes)
'========================================================
Public Sub RunUnlockAll()
    If modSecurity.DEV_MODE Then
        gRunUnlocked = True
        Exit Sub
    End If

    ' Unprotect workbook structure ONCE
    modSecurity.UnprotectWorkbookStructure

    ' Unprotect ALL worksheets ONCE
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
        modSecurity.UnprotectSheet ws

        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            Err.Raise vbObjectError + 9401, , "Sheet still protected after UnprotectSheet: " & ws.name
        End If
    Next ws

    gRunUnlocked = True
End Sub

Public Sub RunReharden()
    If modSecurity.DEV_MODE Then Exit Sub
    On Error Resume Next
    modSecurity.ApplyHardening
    On Error GoTo 0
    gRunUnlocked = False
End Sub
'========================================================
' STEP 1: Pick + validate files (NO writing, NO unlock)
' Returns False when user cancels.
'========================================================
Public Function PickAndValidateFiles(ByRef selectedItemsOut As Variant) As Boolean
    gStep = "PickAndValidateFiles: start"

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = "Select country workbooks to import (MY/SG/CN/JP/...)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx; *.xlsm; *.xls"

        If .Show <> -1 Then
            PickAndValidateFiles = False
            Exit Function
        End If
    End With

    ' Validate ALL selected files first (no import if any fails)
    gStep = "PickAndValidateFiles: ValidateSelectedWorkbooks"
    If Not modImportValidation.ValidateSelectedWorkbooks(fd.selectedItems, gStep) Then
        Err.Raise vbObjectError + 7301, , "Import validation failed. No files imported."
    End If

    Set selectedItemsOut = fd.selectedItems
    PickAndValidateFiles = True
End Function
'========================================================
' STEP 2: Import selected files into tblData (ASSUMES unlocked)
'========================================================
Public Sub ImportSelectedFilesToTblData(ByVal selectedItems As Variant)
    gStep = "ImportSelectedFilesToTblData: start"

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")

    On Error GoTo CleanFail

    gStep = "ImportSelectedFilesToTblData: Ensure schema"
    EnsureTblDataSchema wsData, lo

    Dim keyIndex As Object
    Set keyIndex = BuildExistingKeyIndex(lo)

    Dim i As Long
    For i = 1 To selectedItems.Count
        ImportOneWorkbook CStr(selectedItems(i)), lo, keyIndex
    Next i

    ' Resize table
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastRow < lo.Range.Row Then lastRow = lo.Range.Row
    lo.Resize wsData.Range("A1:L" & lastRow)

    On Error Resume Next
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "mmm-yyyy"
    On Error GoTo 0

    gStep = "ImportSelectedFilesToTblData: done"
    Exit Sub

CleanFail:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------
' Returns True if SlicerCache supports CrossFilterType in this Excel build.
' Excel 2019-safe: uses CallByName so it won't fail compile.
'------------------------------------------------------------
Private Function SupportsSlicerCrossFilterType(ByVal sc As SlicerCache) As Boolean
    On Error GoTo No
    Dim v As Variant
    v = CallByName(sc, "CrossFilterType", VbGet)
    SupportsSlicerCrossFilterType = True
    Exit Function
No:
    SupportsSlicerCrossFilterType = False
End Function

'------------------------------------------------------------
' Best-effort: set slicer to show ONLY items with data.
' Excel 2019-safe: no compile-time enum use, uses numeric constant.
'------------------------------------------------------------
Private Sub SetSlicerShowItemsWithData(ByVal sc As SlicerCache)
    On Error Resume Next
    If SupportsSlicerCrossFilterType(sc) Then
        CallByName sc, "CrossFilterType", VbLet, XL_SLICER_CROSSFILTER_SHOW_ITEMS_WITH_DATA
    End If
    On Error GoTo 0
End Sub
'------------------------------------------------------------
' Build a dictionary of allowed Regions for a given Country from tblData.
' Key = region, Value = True
'------------------------------------------------------------
Private Function GetAllowedRegionsForCountry(ByVal lo As ListObject, ByVal country As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If lo Is Nothing Then
        Set GetAllowedRegionsForCountry = dict
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        Set GetAllowedRegionsForCountry = dict
        Exit Function
    End If

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim c As String, reg As String
        c = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value))   ' Country
        If StrComp(c, country, vbTextCompare) = 0 Then
            reg = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value)) ' Region
            If Len(reg) > 0 Then
                If Not dict.Exists(reg) Then dict.Add reg, True
            End If
        End If
    Next r

    Set GetAllowedRegionsForCountry = dict
End Function

'------------------------------------------------------------
' Build a dictionary of allowed Centers for a given Country from tblData.
' Key = center, Value = True
'------------------------------------------------------------
Private Function GetAllowedCentersForCountry(ByVal lo As ListObject, ByVal country As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If lo Is Nothing Then
        Set GetAllowedCentersForCountry = dict
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        Set GetAllowedCentersForCountry = dict
        Exit Function
    End If

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim c As String, ctr As String
        c = Trim$(CStr(lo.DataBodyRange.Cells(r, 1).Value))   ' Country
        If StrComp(c, country, vbTextCompare) = 0 Then
            ctr = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value)) ' Center
            If Len(ctr) > 0 Then
                If Not dict.Exists(ctr) Then dict.Add ctr, True
            End If
        End If
    Next r

    Set GetAllowedCentersForCountry = dict
End Function

'------------------------------------------------------------
' Hide PivotItems not in allowed dict for a RowField (Region/Center).
' This forces dropdown + slicer list to stop showing other-country items.
'------------------------------------------------------------
Private Sub FilterRowFieldToAllowedItems(ByVal pt As PivotTable, ByVal fieldName As String, ByVal allowed As Object)
    On Error GoTo SafeExit

    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)

    pt.ManualUpdate = True

    ' Avoid Excel auto-adding items back
    On Error Resume Next
    pf.ClearAllFilters
    pf.ShowAllItems = False
    On Error GoTo 0

    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        If LCase$(pi.name) = "(blank)" Then
            pi.Visible = False
        ElseIf allowed.Exists(pi.name) Then
            pi.Visible = True
        Else
            pi.Visible = False
        End If
    Next pi

SafeExit:
    pt.ManualUpdate = False
    On Error Resume Next
    pt.RefreshTable
    DoEvents
    On Error GoTo 0
End Sub

'------------------------------------------------------------
' Apply strict country restriction to Region + Center for given pivot.
'------------------------------------------------------------
'==================================================
' Create PivotCache from tblData but ONLY for 1 country.
' Robust against SpecialCells returning multi-area ranges.
'
' @param lo tblData ListObject
' @param countryFilter e.g. "MY"
' @return PivotCache built from filtered staged range
'==================================================
Private Function CreatePivotCacheFromTable_FilterCountry(ByVal lo As ListObject, ByVal countryFilter As String) As PivotCache
    gStep = "CreatePivotCacheFromTable_FilterCountry: start"

    If lo Is Nothing Then Err.Raise vbObjectError + 3101, , "ListObject is Nothing."
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 3102, , "tblData has no rows."

    NormalizeTblDataHeaders lo

    Dim targetKey As String
    targetKey = NormalizeCountryKey(countryFilter)

    Dim wsHost As Worksheet
    ' Dedicated host per country prevents cross-country cache SourceData overwrite on refresh.
    Set wsHost = EnsureCountryCacheHostSheetVisible(targetKey)

    ' Clear host
    wsHost.Cells.Clear
    wsHost.Cells.ClearFormats
    wsHost.Cells.UnMerge

    Dim colCount As Long
    colCount = lo.Range.Columns.Count
    If colCount < 12 Then colCount = 12 ' hard floor because your schema expects 12

    ' Write headers (force 12 known headers)
    Dim hdr As Variant
    hdr = Array("Country", "Region", "Center", "Date", "Month", "MonthName", _
                "Budget", "Actual", "Variance", "AchievementPct", _
                "UploadId", "UploadedAt")
    wsHost.Range("A1").Resize(1, 12).Value = hdr

    ' Copy matching BODY rows using normalized country key.
    ' This avoids AutoFilter/SpecialCells edge cases and hidden-space mismatches.

    Dim outRow As Long
    outRow = 2 ' start writing at A2

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim rowKey As String
        Dim regionVal As String
        Dim centerVal As String
        Dim monthVal As String
        rowKey = NormalizeCountryKey(lo.DataBodyRange.Cells(r, 1).Value2)
        regionVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 2).Value2))
        centerVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 3).Value2))
        monthVal = Trim$(CStr(lo.DataBodyRange.Cells(r, 5).Value2))

        ' Keep only valid fact rows; prevents (blank) slicer items and malformed pivots.
        If rowKey = targetKey _
           And Len(regionVal) > 0 _
           And Len(centerVal) > 0 _
           And Len(monthVal) > 0 Then
            wsHost.Cells(outRow, 1).Resize(1, 12).Value2 = lo.DataBodyRange.Cells(r, 1).Resize(1, 12).Value2
            wsHost.Cells(outRow, 1).Value2 = targetKey
            outRow = outRow + 1
        End If
    Next r

    Dim dataRows As Long
    dataRows = outRow - 2

    If dataRows <= 0 Then
        Err.Raise vbObjectError + 3104, , "No data rows for country [" & countryFilter & "] after normalized match."
    End If

    ' Build source range on host
    Dim rngDst As Range
    Set rngDst = wsHost.Range("A1").Resize(1 + dataRows, 12)

    ValidatePivotSourceRange rngDst

    Dim srcLocal As String
    srcLocal = rngDst.Address(ReferenceStyle:=xlR1C1, External:=True)

    Dim pc As PivotCache
    On Error GoTo EH
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcLocal)
    pc.MissingItemsLimit = xlMissingItemsNone

    ' Materialize fields from source string
    MaterializePivotCacheFields_FromSourceString srcLocal

    wsHost.Visible = xlSheetVeryHidden

    Set CreatePivotCacheFromTable_FilterCountry = pc
    gStep = "CreatePivotCacheFromTable_FilterCountry: done"
    Exit Function

EH:
    On Error Resume Next
    wsHost.Visible = xlSheetVisible
    On Error GoTo 0

    Err.Raise vbObjectError + 3105, , gStep & ": " & Err.Number & " - " & Err.Description
End Function
'------------------------------------------------------------
' Forces PivotField "Month" to display as mmm-yy in:
' - page fields (filters)
' - row/column label ranges
'
' @param pt Target PivotTable
'------------------------------------------------------------
Private Sub ForceMonthDisplay(ByVal pt As PivotTable)
    On Error GoTo SafeExit

    Dim pf As PivotField
    Set pf = pt.PivotFields("Month")

    pt.ManualUpdate = True
    pt.HasAutoFormat = False

    ' Main: format the pivot field itself (most important for slicers)
    pf.NumberFormat = "mmm-yy"

    ' If Month is a row/column field, format the label cells too
    On Error Resume Next
    If Not pf.LabelRange Is Nothing Then
        pf.LabelRange.NumberFormat = "mmm-yy"
    End If
    On Error GoTo 0

    pt.ManualUpdate = False
    pt.RefreshTable
    DoEvents

SafeExit:
End Sub

'========================================================
' DEBUG: Run this directly to fix chMonth combo on-the-spot
' Also prints diagnostics to Immediate window
'========================================================
Public Sub FixMonthChartCombo()
    ' Manual repair: rebuilds chMonth from correct static range
    On Error GoTo EH
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")
    Dim pt As PivotTable
    Set pt = wsPvt.PivotTables("ptMonth")
    RefreshStaticChartData wsPvt, pt
    Dim shp As Shape
    Set shp = EnsureStaticComboChart(wsDash, wsPvt.Range("AA1").CurrentRegion, _
                "Monthly Budget vs Actual (Overall)", 20, 200, 520, 260, "chMonth")
    MsgBox "Done! HasChart=" & shp.HasChart & " Series=" & _
           IIf(shp.HasChart, shp.Chart.SeriesCollection.Count, 0), vbInformation
    Exit Sub
EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'========================================================
' Resolve a PivotTable data-field caption robustly.
' Handles renamed captions like "Total Budget" vs "Sum of Budget".
'========================================================
Private Function ResolvePivotDataFieldCaption( _
    ByVal pt As PivotTable, _
    ByVal sourceFieldName As String, _
    ByVal preferredCaption As String) As String

    On Error GoTo SafeExit
    If pt Is Nothing Then Exit Function

    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.DataFields(preferredCaption)
    On Error GoTo SafeExit
    If Not pf Is Nothing Then
        ResolvePivotDataFieldCaption = pf.name
        Exit Function
    End If

    On Error Resume Next
    Set pf = pt.DataFields("Sum of " & sourceFieldName)
    On Error GoTo SafeExit
    If Not pf Is Nothing Then
        ResolvePivotDataFieldCaption = pf.name
        Exit Function
    End If

    Dim tok As String
    tok = LCase$(sourceFieldName)

    For Each pf In pt.DataFields
        Dim srcName As String, capName As String
        On Error Resume Next
        srcName = LCase$(CStr(pf.SourceName))
        capName = LCase$(CStr(pf.name))
        On Error GoTo SafeExit

        If srcName = tok Then
            ResolvePivotDataFieldCaption = pf.name
            Exit Function
        End If
        If InStr(1, capName, tok, vbTextCompare) > 0 Then
            ResolvePivotDataFieldCaption = pf.name
            Exit Function
        End If
    Next pf

SafeExit:
End Function

'========================================================
' Refreshes the static data range (Pivots!AA1) used by
' non-PivotCharts. Call this after any pivot refresh.
'========================================================
Public Sub RefreshStaticChartData(ByVal wsPvt As Worksheet, ByVal pt As PivotTable, Optional ByVal forceRefresh As Boolean = False)
    ' Reads Budget/Actual using GetPivotData with robust data-field caption resolution.
    ' Ach% = SUM(Actual)/SUM(Budget) computed per row.
    On Error GoTo SafeExit
    Dim pasteCell As Range
    Set pasteCell = wsPvt.Range("AA1")
    pasteCell.Resize(200, 4).Clear  ' clear fixed 4 cols, not CurrentRegion

    ' Ensure pivot is refreshed and not in manual mode
    If forceRefresh Then
        On Error Resume Next
        pt.ManualUpdate = False
        pt.RefreshTable
        DoEvents
        On Error GoTo SafeExit
    End If

    Dim rf As PivotField
    Set rf = pt.RowFields(1)
    Dim fldBudget As String, fldActual As String
    fldBudget = ResolvePivotDataFieldCaption(pt, "Budget", "Total Budget")
    fldActual = ResolvePivotDataFieldCaption(pt, "Actual", "Total Actual")
    If Len(fldBudget) = 0 Then fldBudget = "Total Budget"
    If Len(fldActual) = 0 Then fldActual = "Total Actual"

    pasteCell.Offset(0, 0).Value = "Month"
    pasteCell.Offset(0, 1).Value = "Total Budget"
    pasteCell.Offset(0, 2).Value = "Total Actual"
    pasteCell.Offset(0, 3).Value = "Ach %"

    Dim iRow As Integer
    iRow = 1
    Dim pi As PivotItem
    For Each pi In rf.PivotItems
        If pi.Visible Then
            Dim dBgt As Double, dAct As Double
            dBgt = 0: dAct = 0
            On Error Resume Next
            ' Try Caption first, then Value (handles date serial vs string)
            dBgt = pt.GetPivotData(fldBudget, rf.name, pi.caption).Value
            If Err.Number <> 0 Then Err.Clear: dBgt = pt.GetPivotData(fldBudget, rf.name, pi.Value).Value
            If Err.Number <> 0 Then dBgt = 0: Err.Clear
            dAct = pt.GetPivotData(fldActual, rf.name, pi.caption).Value
            If Err.Number <> 0 Then Err.Clear: dAct = pt.GetPivotData(fldActual, rf.name, pi.Value).Value
            If Err.Number <> 0 Then dAct = 0: Err.Clear
            On Error GoTo SafeExit
            pasteCell.Offset(iRow, 0).Value = pi.Value
            pasteCell.Offset(iRow, 0).NumberFormat = "mmm-yy"
            pasteCell.Offset(iRow, 1).Value = dBgt
            pasteCell.Offset(iRow, 1).NumberFormat = "#,##0"
            pasteCell.Offset(iRow, 2).Value = dAct
            pasteCell.Offset(iRow, 2).NumberFormat = "#,##0"
            pasteCell.Offset(iRow, 3).Value = IIf(dBgt = 0, 0, dAct / dBgt)
            pasteCell.Offset(iRow, 3).NumberFormat = "0.0%"
            iRow = iRow + 1
        End If
    Next pi
SafeExit:
End Sub

'========================================================
' Refreshes static data range for a given pivot into a
' specified paste cell (e.g. AB1 for ptCountry).
' Same logic as RefreshStaticChartData but configurable.
'========================================================
Public Sub RefreshStaticChartData2(ByVal wsPvt As Worksheet, ByVal pt As PivotTable, ByVal pasteAddr As String, Optional ByVal forceRefresh As Boolean = False)
    On Error GoTo SafeExit
    Dim pasteCell As Range
    Set pasteCell = wsPvt.Range(pasteAddr)
    pasteCell.Resize(200, 4).Clear  ' clear fixed 4 cols, not CurrentRegion

    If forceRefresh Then
        On Error Resume Next
        pt.ManualUpdate = False
        pt.RefreshTable
        DoEvents
        On Error GoTo SafeExit
    End If

    Dim rf As PivotField
    Set rf = pt.RowFields(1)
    Dim fldBudget As String, fldActual As String
    fldBudget = ResolvePivotDataFieldCaption(pt, "Budget", "Total Budget")
    fldActual = ResolvePivotDataFieldCaption(pt, "Actual", "Total Actual")
    If Len(fldBudget) = 0 Then fldBudget = "Total Budget"
    If Len(fldActual) = 0 Then fldActual = "Total Actual"

    pasteCell.Offset(0, 0).Value = rf.name
    pasteCell.Offset(0, 1).Value = "Total Budget"
    pasteCell.Offset(0, 2).Value = "Total Actual"
    pasteCell.Offset(0, 3).Value = "Ach %"

    Dim iRow As Integer
    iRow = 1
    Dim pi As PivotItem
    For Each pi In rf.PivotItems
        If pi.Visible Then
            Dim dBgt As Double, dAct As Double
            dBgt = 0: dAct = 0
            On Error Resume Next
            dBgt = pt.GetPivotData(fldBudget, rf.name, pi.caption).Value
            If Err.Number <> 0 Then Err.Clear: dBgt = pt.GetPivotData(fldBudget, rf.name, pi.Value).Value
            If Err.Number <> 0 Then dBgt = 0: Err.Clear
            dAct = pt.GetPivotData(fldActual, rf.name, pi.caption).Value
            If Err.Number <> 0 Then Err.Clear: dAct = pt.GetPivotData(fldActual, rf.name, pi.Value).Value
            If Err.Number <> 0 Then dAct = 0: Err.Clear
            On Error GoTo SafeExit
            pasteCell.Offset(iRow, 0).Value = pi.Value
            pasteCell.Offset(iRow, 1).Value = dBgt
            pasteCell.Offset(iRow, 1).NumberFormat = "#,##0"
            pasteCell.Offset(iRow, 2).Value = dAct
            pasteCell.Offset(iRow, 2).NumberFormat = "#,##0"
            pasteCell.Offset(iRow, 3).Value = IIf(dBgt = 0, 0, dAct / dBgt)
            pasteCell.Offset(iRow, 3).NumberFormat = "0.0%"
            iRow = iRow + 1
        End If
    Next pi
SafeExit:
End Sub

'========================================================
' Updates an existing non-PivotChart's source data range
' to the refreshed static range, keeping combo formatting.
' chartName = shape name on wsDash (e.g. "chMonth")
' staticAddr = top-left cell of static data (e.g. "AA1")
'========================================================
Public Sub RefreshChartFromStaticRange(ByVal wsDash As Worksheet, ByVal chartName As String, ByVal wsPvt As Worksheet)
    On Error GoTo SafeExit

    Dim shp As Shape
    On Error Resume Next
    Set shp = wsDash.Shapes(chartName)
    On Error GoTo SafeExit
    If shp Is Nothing Then Exit Sub
    If Not shp.HasChart Then Exit Sub

    ' Determine static range address based on chart name
    Dim staticAddr As String
    If chartName = "chMonth" Then
        staticAddr = "AA1"
    ElseIf chartName = "chCountry" Then
        staticAddr = "AH1"
    Else
        Exit Sub
    End If

    Dim staticRng As Range
    Set staticRng = wsPvt.Range(staticAddr).CurrentRegion
    If staticRng.Rows.Count < 2 Then Exit Sub  ' no data rows

    ' Update source data (this keeps series/combo formatting intact)
    shp.Chart.SetSourceData Source:=staticRng

SafeExit:
End Sub

'========================================================
' Called via Application.OnTime after pivot update.
' Runs AFTER Excel finishes PivotChart render - this is
' the only reliable way to change series type on a PivotChart.
'========================================================
Public Sub ApplyAllComboFormats()
    On Error GoTo SafeExit
    Static bRunning As Boolean
    If bRunning Then Exit Sub
    bRunning = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Refresh static ranges + chart data for chMonth and chCountry
    ' (these are plain charts, not PivotCharts - must update manually on slicer change)
    RefreshDashboardCharts

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.HasChart Then
            ApplyComboToChart shp.Chart
        End If
    Next shp
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    bRunning = False
End Sub

'========================================================
' Core combo formatter - works on any chart.
' For PivotCharts: must be called via OnTime (not during
' pivot update) otherwise chartType changes are rejected.
'========================================================
Public Sub ApplyComboToChart(ByVal ch As Chart)
    On Error GoTo SafeExit

    If ch.SeriesCollection.Count = 0 Then Exit Sub

    Dim s As Series
    Dim hasSecondary As Boolean
    hasSecondary = False

    ' Pass 1: axis groups first
    For Each s In ch.SeriesCollection
        If IsAchievementSeries(s) Then
            s.AxisGroup = xlSecondary
            hasSecondary = True
        Else
            s.AxisGroup = xlPrimary
        End If
    Next s

    ' Fallback: force 3rd series as Ach% when name detection is inconsistent.
    If (Not hasSecondary) And ch.SeriesCollection.Count >= 3 Then
        ch.SeriesCollection(3).AxisGroup = xlSecondary
        hasSecondary = True
    End If

    ' Create secondary axis before setting series type
    If hasSecondary Then
        On Error Resume Next
        ch.HasAxis(xlValue, xlSecondary) = True
        On Error GoTo SafeExit
        DoEvents
    End If

    ' Pass 2: set chart type per series
    For Each s In ch.SeriesCollection
        Dim isAch As Boolean
        isAch = IsAchievementSeries(s)
        If (Not isAch) And hasSecondary And ch.SeriesCollection.Count >= 3 Then
            If s.PlotOrder = 3 Then isAch = True
        End If

        On Error Resume Next
        If isAch Then
            s.chartType = xlLineMarkers
            s.AxisGroup = xlSecondary
        Else
            s.chartType = xlColumnClustered
            s.AxisGroup = xlPrimary
        End If
        On Error GoTo SafeExit
    Next s

    ' Pass 3: data labels + number formats + styling
    For Each s In ch.SeriesCollection
        isAch = IsAchievementSeries(s)
        If (Not isAch) And hasSecondary And ch.SeriesCollection.Count >= 3 Then
            If s.PlotOrder = 3 Then isAch = True
        End If
        On Error Resume Next
        s.HasDataLabels = True
        s.DataLabels.ShowValue = True
        If isAch Then
            s.DataLabels.NumberFormat = "0.0%"
            s.DataLabels.Position = xlLabelPositionAbove
            ' Apply gray fill + white text to each label individually.
            ' Must set on the DataLabels collection first, then per label.
            With s.DataLabels
                .Format.Fill.Visible = msoTrue
                .Format.Fill.Solid
                .Format.Fill.ForeColor.RGB = RGB(89, 89, 89)
                .Format.Fill.Transparency = 0
                .Format.Line.Visible = msoFalse
                .Font.Color = RGB(255, 255, 255)
                .Font.size = 9
                .Font.Bold = False
            End With
            ' Also set per individual label (Excel PivotChart override)
            Dim dl As DataLabel
            For Each dl In s.DataLabels
                dl.Format.Fill.Visible = msoTrue
                dl.Format.Fill.Solid
                dl.Format.Fill.ForeColor.RGB = RGB(89, 89, 89)
                dl.Format.Fill.Transparency = 0
                dl.Format.Line.Visible = msoFalse
                dl.Font.Color = RGB(255, 255, 255)
                dl.Font.size = 9
            Next dl
        Else
            s.DataLabels.NumberFormat = "#,##0"
            s.DataLabels.Font.ColorIndex = xlAutomatic
        End If
        On Error GoTo SafeExit
    Next s

    ' Format axes
    On Error Resume Next
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    If hasSecondary Then
        With ch.Axes(xlValue, xlSecondary)
            .TickLabels.NumberFormat = "0%"
            .TickLabelPosition = xlNextToAxis
            .Format.Line.Visible = msoTrue
            .HasTitle = False
        End With
    End If
    On Error GoTo SafeExit

SafeExit:
End Sub

'========================================================
' Programmatically injects Worksheet_PivotTableUpdate
' into the Pivots sheet code module.
' Called at end of every import run so it survives
' sheet rebuilds. Requires "Trust access to VBA project"
' enabled in Excel Trust Center settings.
'========================================================
Public Sub InjectPivotSheetEvents()
    On Error GoTo SafeExit

    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots")

    ' Get the VBComponent for the Pivots sheet
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents(wsPvt.CodeName)

    Dim cm As Object
    Set cm = vbComp.CodeModule

    ' Remove existing Worksheet_PivotTableUpdate if present
    Dim i As Long
    Dim procStart As Long
    On Error Resume Next
    procStart = cm.ProcStartLine("Worksheet_PivotTableUpdate", 0)
    If Err.Number = 0 And procStart > 0 Then
        Dim procLines As Long
        procLines = cm.ProcCountLines("Worksheet_PivotTableUpdate", 0)
        cm.DeleteLines procStart, procLines
    End If
    Err.Clear
    On Error GoTo SafeExit

    ' Build the event code to inject
    Dim code As String
    code = "Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)" & vbCrLf
    code = code & "    Application.OnTime Now + (0.3 / 86400), ""ApplyAllComboFormats""" & vbCrLf
    code = code & "    Application.OnTime Now + (0.8 / 86400), ""ApplyAllComboFormats""" & vbCrLf
    code = code & "    Application.OnTime Now + (2# / 86400), ""ApplyAllComboFormats""" & vbCrLf
    code = code & "End Sub"

    ' Insert at end of module
    Dim lastLine As Long
    lastLine = cm.CountOfLines
    If lastLine = 0 Then
        cm.InsertLines 1, code
    Else
        cm.InsertLines lastLine + 1, code
    End If

SafeExit:
    ' Silently fail if Trust Center blocks VBA extensibility
End Sub

'========================================================
' Deletes known temporary/scratch sheets created as
' side-effects of chart building or pivot operations.
' Called at the very end of EnsurePivotsSlicersChartsAndTables
' so all real charts are fully built before cleanup.
'========================================================
Public Sub DeleteScratchSheets()
    On Error GoTo SafeExit

    ' Names of sheets that should always be deleted if they exist
    Dim scratchNames(2) As String
    scratchNames(0) = "Pivots_Scratch"
    scratchNames(1) = "ChScratch_Tmp"
    scratchNames(2) = "Chart1"

    Application.DisplayAlerts = False

    Dim i As Integer
    For i = 0 To UBound(scratchNames)
        Dim ws As Worksheet
        Dim ch As Chart
        On Error Resume Next

        ' Try as Worksheet first
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(scratchNames(i))
        If Not ws Is Nothing Then
            ws.Delete
            Debug.Print "Deleted worksheet: " & scratchNames(i)
        End If

        ' Try as Chart sheet
        Set ch = Nothing
        Set ch = ThisWorkbook.Charts(scratchNames(i))
        If Not ch Is Nothing Then
            ch.Delete
            Debug.Print "Deleted chart sheet: " & scratchNames(i)
        End If

        On Error GoTo SafeExit
    Next i

    Application.DisplayAlerts = True
    Exit Sub

SafeExit:
    Application.DisplayAlerts = True
End Sub '----------------------------------------------------------------
' BUILD-ONCE, REFRESH-ONLY permanent chart architecture.
'
' First call: creates chMonth/chCountry as plain (non-Pivot) combo
'   chart bound to the static range. Series layout set permanently:
'   col[1]=Budget columns, col[2]=Actual columns, col[3]=Ach% line.
'
' Subsequent calls: chart already exists -> just call SetSourceData
'   with the refreshed range. Series, combo formatting, secondary
'   axis are ALL preserved. No delete, no recreate, no series rebuild.
'
' The static range (AA1 / AH1) always has exactly 4 columns:
'   Label | Total Budget | Total Actual | Ach %
' Rows expand/contract as data changes - chart follows automatically.
'----------------------------------------------------------------

Private Sub UpdateComboChartSeries(ByVal ch As Chart, ByVal srcRng As Range)
    On Error GoTo SafeExit
    If srcRng Is Nothing Or srcRng.Rows.Count < 2 Then GoTo SafeExit
    If ch Is Nothing Then GoTo SafeExit

    Dim nRows As Long: nRows = srcRng.Rows.Count
    LogMsg "[UCS] nRows=" & nRows

    Dim rCat As Range: Set rCat = srcRng.Columns(2).Offset(1, 0).Resize(nRows - 1)
    Dim rBgt As Range: Set rBgt = srcRng.Columns(3).Offset(1, 0).Resize(nRows - 1)
    Dim rAct As Range: Set rAct = srcRng.Columns(4).Offset(1, 0).Resize(nRows - 1)
    Dim rAch As Range: Set rAch = srcRng.Columns(5).Offset(1, 0).Resize(nRows - 1)

    ' Activate chart shape so it is in edit mode
    Dim wsDash As Worksheet
    Set wsDash = ch.Parent.Parent
    Dim wsPrev As Worksheet: Set wsPrev = ActiveSheet
    On Error Resume Next
    wsDash.Activate
    ch.Parent.Select         ' select the shape
    ch.Activate              ' activate the chart (enters edit mode)
    LogMsg "[UCS] chart activated, err=" & Err.Number: Err.Clear

    ' Now assign series data
    If ch.SeriesCollection.Count >= 1 Then
        ch.SeriesCollection(1).XValues = rCat
        ch.SeriesCollection(1).Values = rBgt
        ch.SeriesCollection(1).name = "Total Budget"
        LogMsg "[UCS] s1 err=" & Err.Number: Err.Clear
    End If
    If ch.SeriesCollection.Count >= 2 Then
        ch.SeriesCollection(2).XValues = rCat
        ch.SeriesCollection(2).Values = rAct
        ch.SeriesCollection(2).name = "Total Actual"
        LogMsg "[UCS] s2 err=" & Err.Number: Err.Clear
    End If
    If ch.SeriesCollection.Count >= 3 Then
        ch.SeriesCollection(3).AxisGroup = xlPrimary
        ch.SeriesCollection(3).XValues = rCat
        ch.SeriesCollection(3).Values = rAch
        ch.SeriesCollection(3).name = "Ach %"
        ch.SeriesCollection(3).AxisGroup = xlSecondary
        LogMsg "[UCS] s3 err=" & Err.Number: Err.Clear
    End If

    ' Deactivate chart, restore sheet
    wsDash.Range("A1").Select
    wsPrev.Activate
    On Error GoTo SafeExit
    LogMsg "[UCS] complete"

SafeExit:
    On Error Resume Next
    wsDash.Range("A1").Select
    If Not wsPrev Is Nothing Then wsPrev.Activate
    LogMsg "[UCS] SafeExit err=" & Err.Number & " " & Err.Description
End Sub

Private Function ColLetter(ByVal colNum As Long) As String
    ' Converts column number (1=A, 26=Z, 27=AA...) to Excel letter(s)
    Dim result As String
    Do While colNum > 0
        Dim remainder As Long
        remainder = (colNum - 1) Mod 26
        result = Chr(65 + remainder) & result
        colNum = (colNum - 1) \ 26
    Loop
    ColLetter = result
End Function

Private Function EnsureStaticComboChart( _
    ByVal wsDash As Worksheet, _
    ByVal srcRng As Range, _
    ByVal title As String, _
    ByVal left As Double, _
    ByVal top As Double, _
    ByVal width As Double, _
    ByVal height As Double, _
    ByVal chartName As String _
) As Shape
    ' Full build - called on import. Slicer refresh uses UpdateComboChartData instead.
    Set EnsureStaticComboChart = BuildComboChart(wsDash, srcRng, title, left, top, width, height, chartName)
End Function

Private Function BuildComboChart( _
    ByVal wsDash As Worksheet, _
    ByVal srcRng As Range, _
    ByVal title As String, _
    ByVal left As Double, _
    ByVal top As Double, _
    ByVal width As Double, _
    ByVal height As Double, _
    ByVal chartName As String _
) As Shape
    ' Complete rewrite: build entirely via SERIES formulas.
    ' No SetSourceData, no series trimming, no guessing series count.
    ' 1) AddChart (empty)
    ' 2) Add 3 blank series
    ' 3) Write SERIES formulas directly -> immune to text-column skipping
    ' 4) Format each series explicitly
    On Error Resume Next

    ' Delete old chart
    wsDash.Shapes(chartName).Delete: Err.Clear

    If srcRng Is Nothing Or srcRng.Rows.Count < 2 Then
        LogMsg "[BCH] srcRng empty, abort": Exit Function
    End If

    Dim nRows As Long: nRows = srcRng.Rows.Count
    Dim nD    As Long: nD = nRows - 1      ' data rows (excl header)

    ' Build SERIES formula address strings from srcRng
    ' srcRng layout: col1=Region col2=Center col3=Budget col4=Actual col5=Ach%
    Dim shn As String: shn = "'" & wsDash.name & "'"
    Dim catA As String: catA = shn & "!" & srcRng.Columns(2).Offset(1, 0).Resize(nD).Address(True, True)
    Dim bgtA As String: bgtA = shn & "!" & srcRng.Columns(3).Offset(1, 0).Resize(nD).Address(True, True)
    Dim actA As String: actA = shn & "!" & srcRng.Columns(4).Offset(1, 0).Resize(nD).Address(True, True)
    Dim achA As String: achA = shn & "!" & srcRng.Columns(5).Offset(1, 0).Resize(nD).Address(True, True)
    LogMsg "[BCH] catA=" & catA
    LogMsg "[BCH] bgtA=" & bgtA
    LogMsg "[BCH] achA=" & achA

    ' Create empty chart shape
    Dim shp As Shape
    Set shp = wsDash.Shapes.AddChart(xlColumnClustered, left, top, width, height)
    If shp Is Nothing Then LogMsg "[BCH] AddChart failed": Exit Function
    shp.name = chartName
    LogMsg "[BCH] AddChart err=" & Err.Number: Err.Clear

    Dim ch As Chart: Set ch = shp.Chart

    ' Clear any auto-created series
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
        Err.Clear
    Loop
    LogMsg "[BCH] cleared series err=" & Err.Number: Err.Clear

    ' Add 3 blank series via NewSeries
    ' In OnTime context (slicer refresh), NewSeries sometimes fails due to pivot cache lock.
    ' If it yields < 3 series, fall back to SetSourceData on 5-col range:
    ' Excel auto-skips text cols 1+2 (Region, Center) -> creates exactly 3 numeric series.
    Dim s1 As Series, s2 As Series, s3 As Series
    Set s1 = ch.SeriesCollection.NewSeries: Err.Clear
    Set s2 = ch.SeriesCollection.NewSeries: Err.Clear
    Set s3 = ch.SeriesCollection.NewSeries: Err.Clear
    LogMsg "[BCH] NewSeries count=" & ch.SeriesCollection.Count & " err=" & Err.Number: Err.Clear

    Dim usedFallback As Boolean: usedFallback = False

    If ch.SeriesCollection.Count < 3 Then
        ' NewSeries failed (pivot cache lock in first OnTime call).
        ' SetSourceData works for data but HasAxis(xlSecondary) is permanently
        ' blocked on this chart object - the lock never releases on it.
        ' Strategy: use SetSourceData to show correct bar data NOW (visible immediately),
        ' then queue a Phase-2 OnTime to DELETE this chart and rebuild cleanly via
        ' NewSeries (which works in a second OnTime hop when lock is released).
        LogMsg "[BCH] NewSeries failed, using SetSourceData for immediate display + queuing rebuild"
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
            Err.Clear
        Loop
        ch.SetSourceData Source:=srcRng.Resize(nRows, 5)
        Err.Clear
        LogMsg "[BCH] SetSourceData series count=" & ch.SeriesCollection.Count & " err=" & Err.Number: Err.Clear
        If ch.SeriesCollection.Count < 3 Then
            LogMsg "[BCH] FAIL: SetSourceData also gave < 3 series, aborting"
            Exit Function
        End If
        usedFallback = True

        ' Store all params needed to fully rebuild in Phase-2 OnTime
        gRebuildDashName = wsDash.name
        gRebuildChartName = chartName
        gRebuildTitle = title
        gRebuildLeft = left
        gRebuildTop = top
        gRebuildWidth = width
        gRebuildHeight = height
        ' Derive block address from srcRng top-left cell (col AA = 27)
        gRebuildBlockAddr = "AA" & srcRng.Row
        Application.OnTime Now + (4# / 86400), "Main.RebuildComboChartPhase2"
        LogMsg "[BCH] fallback: phase-2 rebuild queued at +4s for " & chartName
    End If

    If Not usedFallback Then
        ' NewSeries succeeded (import time / no lock) - write SERIES formulas
        ch.SeriesCollection(1).Formula = "=SERIES(""Total Budget""," & catA & "," & bgtA & ",1)"
        LogMsg "[BCH] s1 formula err=" & Err.Number: Err.Clear
        ch.SeriesCollection(2).Formula = "=SERIES(""Total Actual""," & catA & "," & actA & ",2)"
        LogMsg "[BCH] s2 formula err=" & Err.Number: Err.Clear
        ch.SeriesCollection(3).Formula = "=SERIES(""Ach %""," & catA & "," & achA & ",3)"
        LogMsg "[BCH] s3 formula err=" & Err.Number: Err.Clear
    End If

    ' Title
    ch.HasTitle = True: ch.chartTitle.Text = title: Err.Clear

    ' Format s1 + s2: columns, primary, data labels (works in both paths)
    Dim ii As Integer
    For ii = 1 To 2
        With ch.SeriesCollection(ii)
            .chartType = xlColumnClustered
            .AxisGroup = xlPrimary
            .HasDataLabels = True
            .DataLabels.ShowValue = True
            .DataLabels.NumberFormat = "#,##0"
            .DataLabels.Font.size = 8
        End With
        Err.Clear
    Next ii
    LogMsg "[BCH] s1+s2 col+labels err=" & Err.Number: Err.Clear

    If Not usedFallback Then
        ' Not locked - apply secondary axis + s3 immediately
        ApplyComboFormatToChart ch
    End If
    ' usedFallback path: phase-2 OnTime rebuild will apply full formatting

    ' Axis + gap formatting (always safe - not locked)
    On Error Resume Next
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    ch.ChartGroups(1).GapWidth = 80: Err.Clear
    On Error GoTo 0

    Set BuildComboChart = shp
    LogMsg "[BCH] complete series=" & ch.SeriesCollection.Count
End Function

'----------------------------------------------------------------
' Extracts the secondary-axis + s3 combo formatting into a helper.
' Called immediately (non-fallback) or deferred via OnTime (fallback).
'----------------------------------------------------------------
Private Sub ApplyComboFormatToChart(ByVal ch As Chart)
    ' NOTE: When called from ApplyDeferredComboFormat, HasAxis(xlSecondary)=True
    ' has already been set as the lock probe. We skip it here to avoid double-call.
    ' When called from BuildComboChart directly (no fallback / no lock), we set it here.
    On Error Resume Next

    ' Only set secondary axis if not already present (direct call path)
    If Not ch.HasAxis(xlValue, xlSecondary) Then
        ch.HasAxis(xlValue, xlSecondary) = True
        LogMsg "[ACFC] secondary axis created err=" & Err.Number: Err.Clear
    Else
        LogMsg "[ACFC] secondary axis already present"
    End If

    ' Format s3: line on secondary
    With ch.SeriesCollection(3)
        .AxisGroup = xlSecondary
        .chartType = xlLineMarkers
        .HasDataLabels = True
        .DataLabels.ShowValue = True
        .DataLabels.NumberFormat = "0.0%"
        .DataLabels.Position = xlLabelPositionAbove
        .DataLabels.Font.size = 9
    End With
    LogMsg "[ACFC] s3 line+secondary err=" & Err.Number: Err.Clear

    ' Axis number formats
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
    ch.HasAxis(xlCategory, xlSecondary) = False
    AutoScalePercentAxis ch
    LogMsg "[ACFC] complete err=" & Err.Number: Err.Clear
End Sub

'----------------------------------------------------------------
' OnTime callback: applies secondary axis + s3 line formatting
' after pivot cache lock has released (~1.5s after slicer click).
'----------------------------------------------------------------
'----------------------------------------------------------------
' Phase-2 OnTime callback: called ~4s after slicer click.
' By this point we are in a NEW OnTime context - the pivot cache
' lock from the original slicer event is fully released.
' NewSeries works here. We delete the SetSourceData placeholder
' chart and rebuild cleanly with NewSeries + SERIES formulas.
'----------------------------------------------------------------
Public Sub RebuildComboChartPhase2()
    On Error Resume Next
    If Len(gRebuildDashName) = 0 Or Len(gRebuildChartName) = 0 Then Exit Sub

    Dim dashName  As String: dashName = gRebuildDashName
    Dim chartName As String: chartName = gRebuildChartName
    Dim chartTitle As String: chartTitle = gRebuildTitle
    Dim cLeft  As Double: cLeft = gRebuildLeft
    Dim cTop   As Double: cTop = gRebuildTop
    Dim cWidth As Double: cWidth = gRebuildWidth
    Dim cHeight As Double: cHeight = gRebuildHeight
    Dim blockAddr As String: blockAddr = gRebuildBlockAddr

    ' Clear globals immediately to prevent double-fire
    gRebuildDashName = ""
    gRebuildChartName = ""
    gRebuildTitle = ""
    gRebuildBlockAddr = ""

    LogMsg "[P2R] phase-2 rebuild start dash=" & dashName & " chart=" & chartName & " block=" & blockAddr

    Dim wsDash As Worksheet
    Set wsDash = Nothing
    Set wsDash = ThisWorkbook.Worksheets(dashName)
    If wsDash Is Nothing Then LogMsg "[P2R] wsDash not found": Exit Sub

    ' Verify the data block is still valid
    Dim hdrCell As Range
    Set hdrCell = wsDash.Range(blockAddr)
    ' Count actual data rows in block
    Dim iOut As Long: iOut = 0
    Dim r As Long
    For r = 1 To 199
        If Len(Trim$(CStr(hdrCell.Offset(r, 1).Value))) = 0 Then Exit For
        iOut = r
    Next r
    If iOut < 1 Then
        LogMsg "[P2R] no data in block " & blockAddr & ", abort"
        Exit Sub
    End If

    Dim srcRng As Range
    Set srcRng = hdrCell.Resize(iOut + 1, 5)
    LogMsg "[P2R] srcRng=" & srcRng.Address & " iOut=" & iOut

    ' Probe NewSeries - if still locked, retry in 2s
    Dim testCh As Chart
    ' Quick lock test: try NewSeries on a temp basis via the existing chart
    Dim shpExist As Shape
    Set shpExist = wsDash.Shapes(chartName)
    If Not shpExist Is Nothing Then
        If shpExist.HasChart Then
            Dim testS As Series
            Set testS = shpExist.Chart.SeriesCollection.NewSeries
            Dim nsErr As Long: nsErr = Err.Number
            Err.Clear
            If Not testS Is Nothing Then
                testS.Delete: Err.Clear
            End If
            If nsErr <> 0 Or testS Is Nothing Then
                LogMsg "[P2R] NewSeries still locked (err=" & nsErr & "), retrying in 2s"
                gRebuildDashName = dashName
                gRebuildChartName = chartName
                gRebuildTitle = chartTitle
                gRebuildLeft = cLeft
                gRebuildTop = cTop
                gRebuildWidth = cWidth
                gRebuildHeight = cHeight
                gRebuildBlockAddr = blockAddr
                Application.OnTime Now + (2# / 86400), "Main.RebuildComboChartPhase2"
                Exit Sub
            End If
        End If
    End If

    ' Lock released - full clean rebuild via BuildComboChart (NewSeries path will succeed)
    LogMsg "[P2R] lock released, rebuilding " & chartName
    Dim shpNew As Shape
    Set shpNew = BuildComboChart(wsDash, srcRng, chartTitle, cLeft, cTop, cWidth, cHeight, chartName)
    If Not shpNew Is Nothing Then
        ' Re-show the chart (BuildComboChart doesn't set visibility)
        shpNew.Visible = True
    End If
    LogMsg "[P2R] rebuild complete err=" & Err.Number
End Sub

Public Sub UpdateComboChartData( _
    ByVal wsDash As Worksheet, _
    ByVal srcRng As Range, _
    ByVal chartName As String _
)
    ' Strategy 12 revised: rebuild only the visible region chart on every slicer change.
    ' Region slicer -> show different pre-built chart (hide/show, no data change needed).
    ' Month slicer  -> data changes, must rebuild the visible chart with fresh data.
    ' We detect which slicer fired by comparing gPendingSlicerType (set by event handler).
    ' Fallback: always rebuild visible chart data to be safe.
    On Error Resume Next

    Dim country As String
    country = Mid(wsDash.name, Len("Dashboard_") + 1)

    Dim wsPvt As Worksheet
    Set wsPvt = ThisWorkbook.Worksheets("Pivots_" & SanitizeSheetName(country))
    If wsPvt Is Nothing Then Exit Sub
    Dim pt As PivotTable
    Set pt = wsPvt.PivotTables("ptCenter_" & SanitizeName(country))
    If pt Is Nothing Then Exit Sub

    Dim region As String: region = GetSelectedRegions(pt)
    Dim key As String: key = SanitizeName(country)
    Dim prefix As String: prefix = "chCenter_" & key & "_"
    Dim targetName As String: targetName = prefix & region

    LogMsg "[UCD] region=" & region & " target=" & targetName

    ' Determine the correct data block for this region
    Dim blockAddr As String
    blockAddr = GetRegionBlockAddress(pt, region)
    LogMsg "[UCD] blockAddr=" & blockAddr

    ' Write fresh data to that block - use returned iOut for EXPLICIT range
    ' (CurrentRegion bleeds into adjacent 200-row blocks when no blank separator exists)
    Dim iOut As Long
    iOut = RefreshCountryChartData(wsPvt, wsDash, pt, False, blockAddr)

    If iOut < 1 Then
        LogMsg "[UCD] no data in block (iOut=0), exit"
        Exit Sub
    End If

    ' Explicit range: header row + iOut data rows, 5 columns
    Dim srcBlock As Range
    Set srcBlock = wsDash.Range(blockAddr).Resize(iOut + 1, 5)

    ' Delete old chart for this region and rebuild with fresh data
    On Error Resume Next
    Dim shpOld As Shape
    Set shpOld = wsDash.Shapes(targetName)
    If Not shpOld Is Nothing Then shpOld.Delete
    Err.Clear
    On Error Resume Next

    Dim chLeft As Double: chLeft = wsDash.Range("A1").left + 6
    Dim chTop  As Double: chTop = 175
    Dim chH    As Double: chH = 420
    Dim nR As Long: nR = srcBlock.Rows.Count - 1
    Dim chW As Double: chW = Application.Max(400, nR * 55 + 100)

    Dim shpNew As Shape
    Set shpNew = BuildComboChart(wsDash, srcBlock, _
        "Center Performance (" & country & ") - " & IIf(region = "ALL", "All", region), _
        chLeft, chTop, chW, chH, targetName)

    ' Hide all other region charts, show new one
    ShowCountryRegionChart wsDash, country, region
    LogMsg "[UCD] rebuilt & showed " & targetName & " err=" & Err.Number
End Sub

Private Function GetRegionBlockAddress(ByVal pt As PivotTable, ByVal region As String) As String
    ' Returns the helper range start cell for a given region key.
    ' ALL -> AA1, first region -> AA201, second -> AA401, etc.
    ' Must match the order used in BuildCountryRegionCharts.
    If region = "ALL" Then
        GetRegionBlockAddress = "AA1"
        Exit Function
    End If

    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields("Region")
    If pf Is Nothing Then GetRegionBlockAddress = "AA1": Exit Function

    Const BLOCK As Long = 200
    Dim r As Long: r = 0
    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        r = r + 1
        If SanitizeName(pi.name) = region Then
            GetRegionBlockAddress = "AA" & (r * BLOCK + 1)
            Exit Function
        End If
    Next pi

    ' Fallback
    GetRegionBlockAddress = "AA1"
End Function

Private Function GetSelectedRegions(ByVal pt As PivotTable) As String
    ' Returns comma-separated list of selected region items, or "ALL" if all visible
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields("Region")
    If pf Is Nothing Then GetSelectedRegions = "ALL": Exit Function
    On Error Resume Next
    Dim visCount As Long: visCount = 0
    Dim totalCount As Long: totalCount = 0
    Dim lastName As String
    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        totalCount = totalCount + 1
        If pi.Visible Then
            visCount = visCount + 1
            lastName = pi.name
        End If
    Next pi
    If visCount = 0 Or visCount = totalCount Then
        GetSelectedRegions = "ALL"
    ElseIf visCount = 1 Then
        GetSelectedRegions = SanitizeName(lastName)
    Else
        ' Multiple regions - use ALL chart (can't pre-build every combo)
        GetSelectedRegions = "ALL"
    End If
End Function

Public Sub ShowCountryRegionChart(ByVal wsDash As Worksheet, ByVal country As String, ByVal region As String)
    ' Hide all region charts for country, show only the matching one.
    ' Resizes visible chart to fit its actual data point count.
    On Error Resume Next
    Dim key As String: key = SanitizeName(country)
    Dim prefix As String: prefix = "chCenter_" & key & "_"
    Dim targetName As String: targetName = prefix & region

    Dim chLeft As Double: chLeft = wsDash.Range("A1").left + 6
    Dim chTop  As Double: chTop = 175
    Dim chH    As Double: chH = 420

    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.HasChart Then
            If left(shp.name, Len(prefix)) = prefix Then
                If shp.name = targetName Then
                    shp.Visible = True
                    ' Resize to fit data: read point count from series 1
                    Dim nPts As Long: nPts = 20
                    Dim sc As Long: sc = shp.Chart.SeriesCollection.Count
                    If sc >= 1 Then
                        nPts = shp.Chart.SeriesCollection(1).Points.Count
                        If nPts < 1 Then nPts = 20
                    End If
                    Dim chW As Double
                    chW = Application.WorksheetFunction.Max(400, nPts * 55 + 100)
                    shp.left = chLeft
                    shp.top = chTop
                    shp.width = chW
                    shp.height = chH
                    Err.Clear
                Else
                    shp.Visible = False
                End If
            End If
        End If
    Next shp
    LogMsg "[SHW] showed " & targetName & " nPts=" & nPts & " err=" & Err.Number: Err.Clear
End Sub

Public Sub BuildCountryRegionCharts(ByVal wsDash As Worksheet, ByVal wsPvt As Worksheet, _
                                     ByVal pt As PivotTable, ByVal country As String)
    On Error Resume Next
    Dim key As String: key = SanitizeName(country)
    Dim prefix As String: prefix = "chCenter_" & key & "_"
    Const BLOCK As Long = 200

    Dim shp As Shape
    For Each shp In wsDash.Shapes
        If shp.HasChart Then
            If left(shp.name, Len(prefix)) = prefix Then
                shp.Delete: Err.Clear
            End If
        End If
    Next shp
    wsDash.Range("AA1").Resize(BLOCK * 12, 5).Clear: Err.Clear

    Dim pf As PivotField
    Set pf = pt.PivotFields("Region")
    If pf Is Nothing Then LogMsg "[BRC] no Region field": Exit Sub

    Dim regions() As String
    Dim rCount As Long: rCount = 0
    Dim pi As PivotItem
    For Each pi In pf.PivotItems
        ReDim Preserve regions(rCount)
        regions(rCount) = pi.name
        rCount = rCount + 1
    Next pi

    Dim chLeft As Double: chLeft = wsDash.Range("A1").left + 6
    Dim chTop  As Double: chTop = 175
    Dim chH    As Double: chH = 420

    ' ALL chart - block 0 = AA1
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    pt.RefreshTable: Err.Clear
    Application.Wait Now + TimeValue("0:00:01")
    Dim iOutAll As Long
    iOutAll = RefreshCountryChartData(wsPvt, wsDash, pt, False, "AA1")
    If iOutAll < 1 Then
        LogMsg "[BRC] ALL no data, skip"
    Else
        Dim srcAll As Range: Set srcAll = wsDash.Range("AA1").Resize(iOutAll + 1, 5)
        Dim nAll As Long: nAll = iOutAll
        Dim chWAll As Double: chWAll = Application.Max(600, nAll * 55 + 100)
        Dim shpAll As Shape
        Set shpAll = BuildComboChart(wsDash, srcAll, "Center Performance (" & country & ") - All", _
                                      chLeft, chTop, chWAll, chH, prefix & "ALL")
        LogMsg "[BRC] ALL nAll=" & nAll & " err=" & Err.Number: Err.Clear
    End If

    ' Per-region charts
    Dim r As Long
    For r = 0 To rCount - 1
        Dim regName As String: regName = regions(r)
        Dim regKey  As String: regKey = SanitizeName(regName)
        Dim startRow As Long: startRow = (r + 1) * BLOCK + 1
        Dim cellAddr As String: cellAddr = "AA" & startRow

        For Each pi In pf.PivotItems
            pi.Visible = (pi.name = regName)
        Next pi
        pt.RefreshTable: Err.Clear
        Application.Wait Now + TimeValue("0:00:01")

        Dim iOutR As Long
        iOutR = RefreshCountryChartData(wsPvt, wsDash, pt, False, cellAddr)
        If iOutR < 1 Then
            LogMsg "[BRC] skip " & regName & " (no data)"
        Else
            Dim srcR As Range: Set srcR = wsDash.Range(cellAddr).Resize(iOutR + 1, 5)
            Dim nR As Long: nR = iOutR
            Dim chWR As Double: chWR = Application.Max(400, nR * 55 + 100)
            Dim shpR As Shape
            Set shpR = BuildComboChart(wsDash, srcR, _
                "Center Performance (" & country & ") - " & regName, _
                chLeft, chTop, chWR, chH, prefix & regKey)
            If Not shpR Is Nothing Then shpR.Visible = False
            LogMsg "[BRC] " & regName & " nR=" & nR & " err=" & Err.Number: Err.Clear
        End If
    Next r

    ' Restore all regions
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    pt.RefreshTable: Err.Clear

    ShowCountryRegionChart wsDash, country, "ALL"
    LogMsg "[BRC] complete country=" & country
End Sub

Public Sub ApplyDashboardZoom80()
    On Error GoTo SafeExit

    Dim ws As Worksheet
    Dim prevSheet As Worksheet

    Set prevSheet = ActiveSheet
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If LCase$(ws.name) = "dashboard" Or LCase$(left$(ws.name, 10)) = "dashboard_" Then
            ws.Activate
            ActiveWindow.Zoom = 80
        End If
    Next ws

SafeExit:
    On Error Resume Next
    If Not prevSheet Is Nothing Then prevSheet.Activate
    Application.ScreenUpdating = True
End Sub

