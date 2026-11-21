Attribute VB_Name = "modExportPdfChartDirect"
'modExportPdfChartDirect module
Option Explicit

Private Const CS_PREFIX As String = "CS__"
Private gStep As String
Dim generatedAt As Date
'========================================================
' Scratch Pivot Sheet (Option B)
'========================================================

Private Const SCRATCH_SHEET As String = "Pivots_Scratch"
Private Const SCR_PVT_PREFIX As String = "ptScratch_"
Private Const SCR_CH_PREFIX As String = "chScratch_"



Public Sub ExportAllCharts_Fast_ChartSheets_Direct()
    gStep = "Start"
    generatedAt = Now
    
    Dim outPath As String
    Dim newFileName As String
    Dim pdfFolder As String
    Dim archiveFolder As String

    ' Step 1: Ensure pdf\ folder exists
    pdfFolder = GetPdfFolderPath_EnvAware()
    If Not EnsureFolderExists(pdfFolder) Then
        MsgBox "Cannot create/access PDF folder:" & vbCrLf & pdfFolder, vbCritical
        GoTo EH
    End If

    ' Step 2: Ensure pdf\archive\ folder exists
    archiveFolder = pdfFolder & "\archive"
    If Not EnsureFolderExists(archiveFolder) Then
        MsgBox "Cannot create/access archive folder:" & vbCrLf & archiveFolder, vbCritical
        GoTo EH
    End If

    ' Step 3: Build new PDF output path
    newFileName = Format(Now, "yyyy_mm_dd_HHNNSS") & "_" & Format(Timer * 1000 Mod 1000, "000")
    outPath = pdfFolder & "\APAC_Overall_KPI_" & newFileName & ".pdf"

    ' Step 4: Move all existing PDFs from pdf\ into pdf\archive\
    ArchiveExistingPdfs pdfFolder, archiveFolder

    ' Step 5: Clean up archive\ - delete files older than 30 days
    CleanupOldPdfExports archiveFolder, 30, "APAC_Overall_KPI_"

    gStep = "BeginStructureUnlock"
    Main.BeginUnlockedBuild
    RunUnlockAll

    Dim errMsg As String
    Dim errNo As Long
    On Error GoTo EH

    gStep = "App settings: disable"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gStep = "DeleteAllChartSheetsTemp"
    DeleteAllChartSheetsTemp

    Dim exportNames As Collection
    Set exportNames = New Collection

    '=========================
    ' Overall charts (Dashboard)
    '=========================
    gStep = "Add overall charts: chMonth"
    AddChartSheetFromShape exportNames, _
        "APAC Overall Charts - Monthly Budget vs Actual", _
        ThisWorkbook.Worksheets("Dashboard"), "chMonth", generatedAt

    gStep = "Add overall charts: chCountry"
    AddChartSheetFromShape exportNames, _
        "APAC Overall Charts - Country Achievement (Overall)", _
        ThisWorkbook.Worksheets("Dashboard"), "chCountry", generatedAt

    '=========================
    ' Country charts by Region (Scratch Pivot + Scratch Chart)
    '=========================
'=========================
' Country charts by Region (Scratch Pivot + Scratch Chart)
'=========================
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    If LCase$(left$(ws.name, 10)) = "dashboard_" Then

        Dim country As String
        country = Mid$(ws.name, 11)

        Dim wsDash As Worksheet
        Set wsDash = ws

        ' Resolve chart name on dashboard
        Dim chPref As String
        Dim chResolved As String
        chPref = "chCenter_" & SanitizeName_Local(country)

        gStep = "Country [" & country & "]: resolve chart shape " & chPref
        chResolved = ResolveChartShapeName(wsDash, chPref)

        ' Get dashboard chart shape
        Dim shpDashChart As Shape
        gStep = "Country [" & country & "]: get dashboard chart shape " & chResolved
        Set shpDashChart = Nothing
        On Error Resume Next
        Set shpDashChart = wsDash.Shapes(chResolved)
        On Error GoTo EH
        If shpDashChart Is Nothing Then GoTo NextCountry
        If shpDashChart.Type <> msoChart Then GoTo NextCountry

        ' Get original pivot behind dashboard chart (cache + layout reference)
        Dim ptOrig As PivotTable
        gStep = "Country [" & country & "]: get original PivotTable behind dashboard chart"
        Set ptOrig = Nothing
        On Error Resume Next
        Set ptOrig = shpDashChart.Chart.PivotLayout.PivotTable
        DebugPrintRegionOrientation ptOrig, "ORIG before scratch build (" & country & ")"
        On Error GoTo EH
        If ptOrig Is Nothing Then GoTo NextCountry

        Dim ptScratch As PivotTable
        Dim coScratch As ChartObject

        gStep = "Country [" & country & "]: build scratch pivot + chart"
        BuildScratchPivotAndChart ptOrig, country, ptScratch, coScratch

        ' Add Monthly Budget vs Actual page for this country BEFORE region breakdown
        gStep = "Country [" & country & "]: add monthly overview chart page"
        AddCountryMonthlyChartPage exportNames, country, generatedAt

        gStep = "Country [" & country & "]: iterate regions on scratch pivot"
        ExportCountryByRegion_FromScratch exportNames, country, ptScratch, coScratch.Chart, generatedAt
        
    End If

NextCountry:
Next ws

    '=========================
    ' Export all chart sheets to ONE PDF
    '=========================
    gStep = "Export: validate exportNames.Count"
    If exportNames.Count = 0 Then
        MsgBox "No charts found to export." & vbCrLf & _
               "Check chart names and dashboard shapes.", vbExclamation
        GoTo CleanExit
    End If

    ' Prevent overwrite if PDF open
    gStep = "Check PDF lock"
    If Dir(outPath) <> "" Then
        If IsFileLocked(outPath) Then
            MsgBox "Cannot export PDF because it is currently open:" & vbCrLf & vbCrLf & _
                   outPath & vbCrLf & vbCrLf & _
                   "Please close the PDF and try again.", vbExclamation, "PDF In Use"
            GoTo CleanExit
        End If
    End If

    gStep = "Export: ExportSheetNamesToPdf_Safe"
    ExportSheetNamesToPdf_Safe exportNames, outPath

    gStep = "DeleteAllChartSheetsTemp (final)"
    DeleteAllChartSheetsTemp

CleanExit:
    gStep = "CleanExit: restore app settings"
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = False   ' keep suppressed through RunReharden
    RunReharden
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
EH:
    errNo = Err.Number
    errMsg = Err.Description

    If errNo = 0 Then
        errMsg = "Unknown error (Err.Number=0). Most likely a swallowed error before this handler." & vbCrLf & _
                 "Try stepping through the last gStep, or add Debug.Print in each step."
    End If

    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = False   ' ADD
    RunReharden
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "ExportAllCharts_Fast_ChartSheets_Direct failed at:" & vbCrLf & _
           gStep & vbCrLf & _
           errNo & " - " & errMsg, vbCritical
           
   
End Sub
'--------------------------------------------------
' Create a TEMP WORKSHEET and place a PNG of the chart (clipboard-free)
'--------------------------------------------------
Private Sub AddChartSheetFromShape(ByVal names As Collection, ByVal title As String, ByVal srcWs As Worksheet, ByVal shapeName As String, ByVal generatedAt As Date)

    gStep = "AddChartSheetFromShape: lookup shape [" & srcWs.name & "!" & shapeName & "]"

    Dim shp As Shape
    Set shp = Nothing
    On Error Resume Next
    Set shp = srcWs.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    If shp.Type <> msoChart Then Exit Sub

    gStep = "AddChartSheetFromShape: export chart to PNG"
    Dim tmpPng As String
    tmpPng = GetTempPngPath(CS_PREFIX & Format$(names.Count + 1, "000"))

    ' Ensure chart is ready, then export (no clipboard)
    On Error Resume Next
    shp.Chart.Refresh
    On Error GoTo 0
    DoEvents

    If Not shp.Chart.Export(fileName:=tmpPng, FilterName:="PNG") Then
        Err.Raise vbObjectError + 2201, , "Chart.Export failed for [" & srcWs.name & "!" & shapeName & "]"
    End If

    gStep = "AddChartSheetFromShape: create temp worksheet"
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.name = CS_PREFIX & Format$(names.Count + 1, "000")

    ' Clear and layout
    wsOut.Cells.Clear
    wsOut.Range("A1").Value = title
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.size = 14

    gStep = "AddChartSheetFromShape: place PNG"
    Dim piC As Shape
    Set piC = wsOut.Shapes.AddPicture( _
        fileName:=tmpPng, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        left:=wsOut.Range("A3").left, _
        top:=wsOut.Range("A3").top, _
        width:=-1, _
        height:=-1 _
    )

    ' Fit picture to page nicely (landscape)
    gStep = "AddChartSheetFromShape: fit picture"
    FitShapeToPage wsOut, piC, wsOut.Range("A3").left, wsOut.Range("A3").top

    ' Page setup (safe)
    gStep = "AddChartSheetFromShape: pagesetup"
    ' Page setup (worksheet is SAFE to use Zoom=False)
    With wsOut.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    gStep = "AddChartSheetFromShape: footer"
    ApplyPdfFooter wsOut, generatedAt, "", ThisWorkbook.name

    ' Tight print area: from A1 to picture bottom-right cell
    Dim br As Range
    Set br = piC.BottomRightCell
    
    wsOut.PageSetup.PrintArea = wsOut.Range("A1", br).Address

    names.Add wsOut.name

End Sub

'--------------------------------------------------
' Safer PDF export (no Select)
'--------------------------------------------------
Private Sub ExportSheetNamesToPdf_Safe(ByVal names As Collection, ByVal outPath As String)

    gStep = "ExportSheetNamesToPdf_Safe: build sheets array"

    Dim arr() As Variant
    Dim i As Long
    ReDim arr(1 To names.Count)
    For i = 1 To names.Count
        arr(i) = names(i)
    Next i

    gStep = "ExportSheetNamesToPdf_Safe: select sheets"
    ThisWorkbook.Worksheets(arr).Select

    gStep = "ExportSheetNamesToPdf_Safe: export"
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=outPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    gStep = "ExportSheetNamesToPdf_Safe: return to Dashboard"
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    On Error GoTo 0

End Sub

Private Sub DeleteAllChartSheetsTemp()

    gStep = "DeleteAllChartSheetsTemp: start"

    Dim wasProtected As Boolean
    'wasProtected = BeginStructureUnlock_Local("DeleteAllChartSheetsTemp")

    On Error Resume Next
    Application.DisplayAlerts = False

    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If LCase$(left$(ThisWorkbook.Worksheets(i).name, Len(CS_PREFIX))) = LCase$(CS_PREFIX) Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next i

    Application.DisplayAlerts = True
    'EndStructureUnlock_Local wasProtected
    On Error GoTo 0

    gStep = "DeleteAllChartSheetsTemp: done"
End Sub


Private Function BeginStructureUnlock(ByVal actionName As String) As Boolean
    If modSecurity.DEV_MODE Then
        BeginStructureUnlock = False
        Exit Function
    End If

    On Error GoTo EH

    BeginStructureUnlock = ThisWorkbook.ProtectStructure
    If BeginStructureUnlock Then
        ThisWorkbook.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
    End If
    Exit Function

EH:
    ' if anything goes wrong, treat as "not protected"
    BeginStructureUnlock = False
End Function




'========================================================
' Pivot helpers
'========================================================
Private Function FindPivotByPrefix(ByVal ws As Worksheet, ByVal prefix As String) As PivotTable
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        If LCase$(left$(pt.name, Len(prefix))) = LCase$(prefix) Then
            Set FindPivotByPrefix = pt
            Exit Function
        End If
    Next pt
    Set FindPivotByPrefix = Nothing
End Function

Private Function SanitizeName_Local(ByVal s As String) As String
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
    SanitizeName_Local = t
End Function

Private Function ResolveChartShapeName(ByVal ws As Worksheet, ByVal preferredName As String) As String
    On Error Resume Next

    Dim shp As Shape

    Set shp = ws.Shapes(preferredName)
    If Not shp Is Nothing Then
        ResolveChartShapeName = preferredName
        Exit Function
    End If

    Dim nm As String, prefL As String
    prefL = LCase$(preferredName)

    For Each shp In ws.Shapes
        nm = LCase$(shp.name)
        If left$(nm, Len(prefL)) = prefL Then
            ResolveChartShapeName = shp.name
            Exit Function
        End If
    Next shp

    For Each shp In ws.Shapes
        If shp.Type = msoChart Then
            ResolveChartShapeName = shp.name
            Exit Function
        End If
    Next shp

    ResolveChartShapeName = preferredName
End Function

'--------------------------------------------------
' Select exactly ONE slicer item by toggling item visibility.
' Works when VisibleSlicerItemsList throws 1004.
'
' @param sc The slicer cache
' @param itemName The slicer item caption/name to select
'--------------------------------------------------
Private Sub ApplySingleSlicerSelection(ByVal sc As SlicerCache, ByVal itemName As String)

    gStep = "ApplySingleSlicerSelection: " & sc.name & " -> " & itemName

    On Error GoTo EH

    ' Make sure cache is ready
    DoEvents

    ' Clear filter first
    sc.ClearManualFilter
    DoEvents

    Dim si As SlicerItem

    ' First pass: set all items unselected (Visible=False)
    For Each si In sc.SlicerItems
        If si.name <> "(blank)" Then
            On Error Resume Next
            si.Selected = False
            On Error GoTo EH
        End If
    Next si

    ' Second pass: select the target
    For Each si In sc.SlicerItems
        If StrComp(si.name, itemName, vbTextCompare) = 0 Then
            si.Selected = True
            Exit Sub
        End If
    Next si

    ' If item not found, fail fast
    Err.Raise vbObjectError + 2401, , "Slicer item not found in [" & sc.name & "]: [" & itemName & "]"

EH:
    Err.Raise vbObjectError + 2001, , _
        "ApplySingleSlicerSelection failed for [" & sc.name & "] item [" & itemName & "]: " & Err.Number & " - " & Err.Description
End Sub

Private Sub ClearSlicerFilter(ByVal sc As SlicerCache)
    gStep = "ClearSlicerFilter: " & sc.name
    On Error Resume Next
    sc.ClearManualFilter
    On Error GoTo 0
End Sub



'--------------------------------------------------
' Copies a chart into a ChartSheet reliably.
' Avoids ChartArea.Copy (unstable for PivotCharts).
'
' @param srcChart The source Chart (from Shape.Chart)
' @param dstSheet The destination Chart sheet
'--------------------------------------------------
Private Sub CopyChartToChartSheet(ByVal srcChart As Chart, ByVal dstSheet As Chart)

    Dim attempt As Long

    For attempt = 1 To 5

        On Error Resume Next
        Application.CutCopyMode = False
        Err.Clear
        On Error GoTo 0

        ' Ensure chart is up-to-date before copying
        On Error Resume Next
        srcChart.Refresh
        On Error GoTo 0
        DoEvents

        ' Prefer Chart.Copy (much more stable than ChartArea.Copy)
        On Error Resume Next
        srcChart.Copy
        If Err.Number = 0 Then
            Err.Clear
            dstSheet.Activate
            DoEvents

            dstSheet.Paste
            If Err.Number = 0 Then
                Exit Sub
            End If
        End If

        ' If paste failed, wait and retry
        On Error Resume Next
        Err.Clear
        Application.Wait Now + TimeValue("0:00:01")
        On Error GoTo 0
        DoEvents
    Next attempt

    ' If we reached here, it consistently failed
    Err.Raise vbObjectError + 2101, , _
        "CopyChartToChartSheet failed after retries. " & _
        "Common causes: PivotChart still refreshing, clipboard blocked, or chart is hidden/off-screen."
End Sub

'--------------------------------------------------
' Build a temp file path for PNG export
'--------------------------------------------------
Private Function GetTempPngPath(ByVal baseName As String) As String
    Dim tmpFolder As String
    tmpFolder = Environ$("TEMP")
    If Len(tmpFolder) = 0 Then tmpFolder = Environ$("USERPROFILE") & "\AppData\Local\Temp"
    GetTempPngPath = tmpFolder & "\" & baseName & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".png"
End Function

'--------------------------------------------------
' Resize a shape to fit within the printable page width/height area
' (simple, reliable "good enough" fit)
'--------------------------------------------------
Private Sub FitShapeToPage(ByVal ws As Worksheet, ByVal shp As Shape, ByVal leftPos As Double, ByVal topPos As Double)

    ' Conservative max size (tuned for typical A4/Letter landscape)
    ' You can adjust these if you want bigger.
    Dim maxW As Double: maxW = 900
    Dim maxH As Double: maxH = 480

    shp.LockAspectRatio = msoTrue

    ' Scale down if needed
    If shp.width > maxW Then shp.width = maxW
    If shp.height > maxH Then shp.height = maxH

    ' After width change, height changes automatically due to aspect ratio
    If shp.height > maxH Then shp.height = maxH

    shp.left = leftPos
    shp.top = topPos
End Sub

'--------------------------------------------------
' Filters a PivotField (row field OK) by caption equals.
' Reliable for export.
'--------------------------------------------------
'--------------------------------------------------
' Filters a PivotField to exactly one item (by caption).
' Works for:
' - PageFields (Report Filters): uses CurrentPage
' - Row/Column fields: uses PivotFilters (caption equals)
'
' @param pt Target PivotTable
' @param fieldName Pivot field name (e.g. "Region")
' @param caption Pivot item caption to select (e.g. "AH")
'--------------------------------------------------
Private Sub ApplyPivotFieldEquals(ByVal pt As PivotTable, ByVal fieldName As String, ByVal caption As String)

    gStep = "ApplyPivotFieldEquals: " & pt.name & " " & fieldName & "=" & caption

    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName) ' will throw if field missing

    ' Always start clean
    On Error Resume Next
    pf.ClearAllFilters
    On Error GoTo 0

    ' PAGE FIELD (Report Filter)
    If pf.Orientation = xlPageField Then
        ' Ensure single selection
        On Error Resume Next
        pf.EnableMultiplePageItems = False
        On Error GoTo 0

        ' Validate item exists (prevents cryptic errors)
        Dim pi As PivotItem
        On Error Resume Next
        Set pi = pf.PivotItems(caption)
        On Error GoTo 0
        If pi Is Nothing Then
            Err.Raise vbObjectError + 5401, , _
                "Pivot item not found for PageField [" & fieldName & "]: [" & caption & "] in " & pt.name
        End If

        pf.CurrentPage = caption
        Exit Sub
    End If

    ' ROW/COLUMN FIELD
    ' Remove existing PivotFilters (important)
    On Error Resume Next
    Dim i As Long
    For i = pf.PivotFilters.Count To 1 Step -1
        pf.PivotFilters(i).Delete
    Next i
    On Error GoTo 0

    ' Apply caption filter
    pf.PivotFilters.Add2 Type:=xlCaptionEquals, Value1:=caption

End Sub
Private Sub ClearPivotFieldFilters(ByVal pt As PivotTable, ByVal fieldName As String)

    gStep = "ClearPivotFieldFilters: " & pt.name & " " & fieldName

    Dim pf As PivotField
    Set pf = pt.PivotFields(fieldName)

    On Error Resume Next
    pf.ClearAllFilters
    On Error GoTo 0

    If pf.Orientation = xlPageField Then
        ' Return to (All) if possible
        On Error Resume Next
        pf.CurrentPage = "(All)"
        On Error GoTo 0
        Exit Sub
    End If

    ' Row/column: remove pivot filters
    On Error Resume Next
    Dim i As Long
    For i = pf.PivotFilters.Count To 1 Step -1
        pf.PivotFilters(i).Delete
    Next i
    On Error GoTo 0

End Sub
'--------------------------------------------------
' Applies a standard footer to a temp export sheet.
' Uses Excel header/footer codes so it prints in PDF.
'
' @param wsOut Temp chart worksheet (CS__###)
' @param generatedAt The timestamp you want to show (one value for whole export)
' @param leftText Optional left footer text
' @param rightText Optional right footer text
'--------------------------------------------------
Private Sub ApplyPdfFooter(ByVal wsOut As Worksheet, ByVal generatedAt As Date, _
                           Optional ByVal leftText As String = "", _
                           Optional ByVal rightText As String = "")

    gStep = "ApplyPdfFooter: " & wsOut.name

    Dim stamp As String
    stamp = "Generated: " & Format$(generatedAt, "yyyy-mm-dd HH:nn:ss")

    ' You can tune font size by adding: &"Calibri,8"
    ' Example: "&""Calibri,8""Generated: ..."
    With wsOut.PageSetup
        .LeftFooter = leftText
        .CenterFooter = stamp
        .RightFooter = "Page &P of &N"
    End With
    gStep = "AddChartSheetFromShape: footer"
    'ApplyPdfFooter wsOut, generatedAt, "", ThisWorkbook.name
End Sub

'========================================================
' Ensure a folder exists (creates if missing)
'
' @param folderPath Full absolute folder path
' @return True if folder ready, False if failed
'========================================================
Public Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(folderPath) = 0 Then Exit Function

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    EnsureFolderExists = True
    Exit Function

EH:
    EnsureFolderExists = False
End Function
'========================================================
' Delete old exported PDFs in a folder (keep last N days)
'
' Rules:
' 1) Only deletes *.pdf
' 2) Only deletes files whose names start with filePrefix (optional)
' 3) Only deletes if file is older than keepDays
'
' @param folderPath Folder that contains the exports (e.g., ThisWorkbook.Path & "\pdf")
' @param keepDays Number of days to keep (e.g., 30)
' @param filePrefix Optional prefix filter (e.g., "APAC_Overall_KPI_"). "" = no prefix filter
' @return Number of files deleted
'========================================================
Public Function CleanupOldPdfExports( _
    ByVal folderPath As String, _
    ByVal keepDays As Long, _
    Optional ByVal filePrefix As String = "APAC_Overall_KPI_" _
) As Long

    On Error GoTo EH

    Dim fso As Object
    Dim folder As Object
    Dim f As Object
    Dim cutoff As Date
    Dim deletedCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        CleanupOldPdfExports = 0
        Exit Function
    End If

    Set folder = fso.GetFolder(folderPath)

    cutoff = DateAdd("d", -keepDays, Now)

    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.name)) = "pdf" Then

            ' Optional prefix filter
            If Len(filePrefix) = 0 Or left$(f.name, Len(filePrefix)) = filePrefix Then

                ' Use DateLastModified; you can change to DateCreated if you prefer
                If f.DateLastModified < cutoff Then
                    ' Avoid crash if file is read-only or locked: attempt delete, ignore failure
                    On Error Resume Next
                    f.Delete True
                    If Err.Number = 0 Then
                        deletedCount = deletedCount + 1
                    End If
                    Err.Clear
                    On Error GoTo EH
                End If

            End If
        End If
    Next f

    CleanupOldPdfExports = deletedCount
    Exit Function

EH:
    ' Fail safe: return what we deleted so far
    CleanupOldPdfExports = deletedCount
End Function
'========================================================
' Move all *.pdf files in pdfFolder into archiveFolder.
' Skips files that are locked/open (silently).
' If destination filename already exists, appends _1, _2 etc.
'
' @param pdfFolder     Source folder  e.g. "...\pdf"
' @param archiveFolder Destination    e.g. "...\pdf\archive"
'========================================================
Public Sub ArchiveExistingPdfs( _
    ByVal pdfFolder As String, _
    ByVal archiveFolder As String)

    On Error GoTo EH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(pdfFolder) Then Exit Sub

    Dim oFolder As Object
    Set oFolder = fso.GetFolder(pdfFolder)

    Dim f As Object
    Dim destPath As String
    Dim baseName As String
    Dim ext As String
    Dim counter As Long

    For Each f In oFolder.Files
        If LCase$(fso.GetExtensionName(f.name)) = "pdf" Then

            baseName = fso.GetBaseName(f.name)
            ext = "." & fso.GetExtensionName(f.name)
            destPath = archiveFolder & "\" & f.name

            ' Resolve filename collision in archive
            counter = 0
            Do While fso.FileExists(destPath)
                counter = counter + 1
                destPath = archiveFolder & "\" & baseName & "_" & counter & ext
            Loop

            ' Move - skip silently if file is locked or in use
            On Error Resume Next
            fso.MoveFile f.path, destPath
            If Err.Number <> 0 Then
                Debug.Print "ArchiveExistingPdfs: skipped locked file [" & f.name & "] - " & Err.Description
                Err.Clear
            End If
            On Error GoTo EH
        End If
    Next f

    Exit Sub
EH:
    Debug.Print "ArchiveExistingPdfs error: " & Err.Number & " - " & Err.Description
End Sub

Private Function GetPdfFolderPath_EnvAware() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.path

    ' Normalize slashes
    wbPath = Replace(wbPath, "/", "\")

    If Len(wbPath) = 0 Or _
       InStr(1, wbPath, "https:\\", vbTextCompare) > 0 Or _
       InStr(1, wbPath, "http:\\", vbTextCompare) > 0 Then

        ' Opened from SharePoint URL — map to local OneDrive path
        ' SharePoint: https://village365apc-my.sharepoint.com/personal/haishi_tan_davita_com/Documents/Malaysia/Projects/APAC_KPI
        ' Maps to:    C:\Users\HTAN\OneDrive - DaVita\Malaysia\Projects\APAC_KPI
        Dim oneDriveRoot As String
        oneDriveRoot = Environ$("OneDriveCommercial")
        If Len(oneDriveRoot) = 0 Then oneDriveRoot = Environ$("OneDrive")
        If Len(oneDriveRoot) = 0 Then
            Err.Raise vbObjectError + 8711, , "OneDrive environment variable not found."
        End If

        ' Extract the path after /Documents/ from the SharePoint URL
        ' e.g. "Malaysia/Projects/APAC_KPI" -> "Malaysia\Projects\APAC_KPI"
        Dim marker As String
        marker = "Malaysia_IT\"
        Dim markerPos As Long
        markerPos = InStr(1, wbPath, marker, vbTextCompare)

        Dim subPath As String
        If markerPos > 0 Then
            'subPath = Mid(wbPath, markerPos + Len(marker))
            subPath = "Malaysia Information Technology - APAC KPI"
        Else
            subPath = ""
        End If

        If Len(subPath) > 0 Then
            GetPdfFolderPath_EnvAware = oneDriveRoot & "\" & subPath & "\pdf"
        Else
            ' Fallback if URL format unexpected
            GetPdfFolderPath_EnvAware = oneDriveRoot & "\Malaysia\Projects\APAC_KPI\pdf"
        End If

    Else
        ' Local or OneDrive-synced local path — save pdf\ beside the workbook
        GetPdfFolderPath_EnvAware = wbPath & "\pdf"
    End If

    Debug.Print "PDF folder resolved: " & GetPdfFolderPath_EnvAware
End Function

'========================================================
' DEV/PROD PDF folder resolver
' DEV (SharePoint URL workbook) -> C:\Users\HTAN\OneDrive - DaVita\Malaysia\Projects\APAC_KPI\pdf
' PROD (local workbook)         -> C:\Users\HTAN\OneDrive - DaVita\Malaysia Information Technology - APAC KPI\pdf
'========================================================
Private Function GetDefaultPdfPath_EnvAware(ByVal fileName As String) As String
    Dim pdfFolder As String
    pdfFolder = GetPdfFolderPath_EnvAware()

    ' Create folder (FSO CreateFolder only creates one level; your paths already exist up to root)
    If Not EnsureFolderExists(pdfFolder) Then
        Err.Raise vbObjectError + 8712, , "Cannot create/access PDF folder:" & vbCrLf & pdfFolder
    End If

    GetDefaultPdfPath_EnvAware = pdfFolder & "\" & fileName
End Function

'--------------------------------------------------
' Build scratch pivot + chart for one country.
' Writes country-filtered data to scratch sheet.
' Chart is built from a STATIC range (not pivot source)
' so NewSeries works without pivot-binding conflicts.
' Ach% = SUM(Actual)/SUM(Budget) per Center, computed
' directly from the static data - no CalculatedFields needed.
'--------------------------------------------------
Private Sub BuildScratchPivotAndChart( _
    ByVal ptOrig As PivotTable, _
    ByVal country As String, _
    ByRef ptScratch As PivotTable, _
    ByRef coScratch As ChartObject)

    Dim safeCountry As String
    safeCountry = SanitizeName_Local(country)


    gStep = "BuildScratchPivotAndChart: hard-recreate scratch sheet for " & safeCountry
    Dim wsScratch As Worksheet
    Set wsScratch = CreateFreshScratchPivotSheet()
    wsScratch.Visible = xlSheetVisible

    '==============================================
    ' Step 1: Find source columns in tblData
    '==============================================
    gStep = "BuildScratchPivotAndChart: locate tblData for " & country

    Dim loSrc As ListObject
    On Error Resume Next
    Set loSrc = ThisWorkbook.Worksheets("Data").ListObjects("tblData")
    On Error GoTo EH
    If loSrc Is Nothing Then
        Err.Raise vbObjectError + 9201, , "tblData not found on Data sheet"
    End If

    Dim srcBody As Range
    Set srcBody = loSrc.DataBodyRange
    If srcBody Is Nothing Then
        Err.Raise vbObjectError + 9202, , "tblData has no data rows"
    End If

    Dim colCountry As Long, colRegion As Long, colCenter As Long
    Dim colBudget As Long, colActual As Long
    Dim hdr As ListColumn
    For Each hdr In loSrc.ListColumns
        Select Case LCase$(hdr.name)
            Case "country": colCountry = hdr.Index
            Case "region":  colRegion = hdr.Index
            Case "center":  colCenter = hdr.Index
            Case "budget":  colBudget = hdr.Index
            Case "actual":  colActual = hdr.Index
        End Select
    Next hdr
    If colCountry = 0 Or colRegion = 0 Or colCenter = 0 Or colBudget = 0 Or colActual = 0 Then
        Err.Raise vbObjectError + 9203, , "Required columns not found in tblData"
    End If

    '==============================================
    ' Step 2: Write country-only data to scratch sheet
    ' A=Region  B=Center  C=Budget  D=Actual
    '==============================================
    gStep = "BuildScratchPivotAndChart: write filtered data for " & country

    wsScratch.Range("A1:D1").Value = Array("Region", "Center", "Budget", "Actual")

    Dim outRow As Long
    outRow = 2
    Dim r As Long
    Dim bVal As Double, aVal As Double
    For r = 1 To srcBody.Rows.Count
        If StrComp(Trim$(CStr(srcBody.Cells(r, colCountry).Value)), country, vbTextCompare) = 0 Then
            bVal = 0: aVal = 0
            On Error Resume Next
            bVal = CDbl(srcBody.Cells(r, colBudget).Value)
            aVal = CDbl(srcBody.Cells(r, colActual).Value)
            On Error GoTo EH
            wsScratch.Cells(outRow, 1).Value = srcBody.Cells(r, colRegion).Value
            wsScratch.Cells(outRow, 2).Value = srcBody.Cells(r, colCenter).Value
            wsScratch.Cells(outRow, 3).Value = bVal
            wsScratch.Cells(outRow, 4).Value = aVal
            outRow = outRow + 1
        End If
    Next r

    Dim dataRows As Long
    dataRows = outRow - 2
    If dataRows = 0 Then
        Err.Raise vbObjectError + 9204, , "No data found for country [" & country & "]"
    End If

    '==============================================
    ' Step 3: Build PivotCache + PivotTable
    ' Layout: Rows=Center, PageField=Region
    ' Values: Budget SUM, Actual SUM  (only 2 fields - no Ach% in pivot)
    '==============================================
    gStep = "BuildScratchPivotAndChart: build PivotCache"
    Dim srcRange As Range
    Set srcRange = wsScratch.Range("A1").Resize(1 + dataRows, 4)

    Dim freshCache As PivotCache
    Set freshCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)
    freshCache.MissingItemsLimit = xlMissingItemsNone

    ' Place pivot well below the raw data
    Dim pvtAnchor As Range
    Set pvtAnchor = wsScratch.Cells(dataRows + 4, 1)

    Dim ptName As String
    ptName = SCR_PVT_PREFIX & safeCountry

    gStep = "BuildScratchPivotAndChart: CreatePivotTable " & ptName
    Set ptScratch = freshCache.CreatePivotTable( _
        TableDestination:=pvtAnchor, _
        TableName:=ptName)

    gStep = "BuildScratchPivotAndChart: build pivot layout"
    ptScratch.ManualUpdate = True
    ptScratch.AddFields RowFields:="Center", PageFields:="Region"

    Dim fBudget As PivotField
    Set fBudget = ptScratch.AddDataField( _
        ptScratch.PivotFields("Budget"), "Total Budget", xlSum)
    fBudget.NumberFormat = "#,##0"

    Dim fActual As PivotField
    Set fActual = ptScratch.AddDataField( _
        ptScratch.PivotFields("Actual"), "Total Actual", xlSum)
    fActual.NumberFormat = "#,##0"

    ' Remove grand total row — must be set before refresh
    ptScratch.RowGrand = False
    ptScratch.ColumnGrand = False

    ptScratch.ManualUpdate = False
    ptScratch.RefreshTable
    DoEvents

    '==============================================
    ' Step 4: Build STATIC chart data block
    ' This is KEY: chart is NOT bound to pivot.
    ' We read pivot DataBodyRange into a static block,
    ' compute Ach% = Actual/Budget, then build all
    ' 3 series from plain ranges.
    ' Column layout (starting col F):
    '   F=Center label  G=Total Budget  H=Total Actual  I=Ach%
    '==============================================
    gStep = "BuildScratchPivotAndChart: build static chart data block"

    Dim pvtBody As Range
    Set pvtBody = ptScratch.DataBodyRange   ' rows=centers, col1=Budget, col2=Actual

    Dim centerCount As Long
    centerCount = pvtBody.Rows.Count

    ' Static block starts at F1
    Dim blkCol As Long
    blkCol = 6  ' col F
    Dim blkRow As Long
    blkRow = 1  ' header at row 1

    ' Headers
    wsScratch.Cells(blkRow, blkCol).Value = "Center"
    wsScratch.Cells(blkRow, blkCol + 1).Value = "Total Budget"
    wsScratch.Cells(blkRow, blkCol + 2).Value = "Total Actual"
    wsScratch.Cells(blkRow, blkCol + 3).Value = "Ach %"

    ' Read pivot row labels — RowRange includes header + data rows
    ' Row labels are in ptScratch.RowRange, data rows start at index 2
    Dim rowLabelRange As Range
    Set rowLabelRange = ptScratch.rowRange  ' col of center labels incl header

    Dim i As Long
    For i = 1 To centerCount
        Dim centerLabel As String
        centerLabel = ""
        On Error Resume Next
        centerLabel = CStr(rowLabelRange.Cells(i + 1, 1).Value) ' +1 = skip header row
        On Error GoTo EH

        bVal = 0: aVal = 0
        On Error Resume Next
        bVal = CDbl(pvtBody.Cells(i, 1).Value)  ' Budget col
        aVal = CDbl(pvtBody.Cells(i, 2).Value)  ' Actual col
        On Error GoTo EH

        wsScratch.Cells(blkRow + i, blkCol).Value = centerLabel
        wsScratch.Cells(blkRow + i, blkCol + 1).Value = bVal
        wsScratch.Cells(blkRow + i, blkCol + 2).Value = aVal
        If bVal <> 0 Then
            wsScratch.Cells(blkRow + i, blkCol + 3).Value = aVal / bVal
        Else
            wsScratch.Cells(blkRow + i, blkCol + 3).Value = 0
        End If
    Next i

    ' Define the 4 static ranges
    Dim rLabels  As Range
    Dim rBudget  As Range
    Dim rActual  As Range
    Dim rAchPct  As Range
    Set rLabels = wsScratch.Cells(blkRow + 1, blkCol).Resize(centerCount, 1)
    Set rBudget = wsScratch.Cells(blkRow + 1, blkCol + 1).Resize(centerCount, 1)
    Set rActual = wsScratch.Cells(blkRow + 1, blkCol + 2).Resize(centerCount, 1)
    Set rAchPct = wsScratch.Cells(blkRow + 1, blkCol + 3).Resize(centerCount, 1)
    
    '==============================================
    ' Step 5: Build chart on a SEPARATE temp sheet
    ' (no pivot on that sheet = plain chart = NewSeries works)
    '==============================================
    Dim chName As String
    chName = SCR_CH_PREFIX & safeCountry

    ' Create a separate plain sheet for the chart only
    gStep = "BuildScratchPivotAndChart: create chart-only sheet"
    Dim wsChart As Worksheet
    Dim chartSheetName As String
    chartSheetName = "ChScratch_Tmp"

    ' Must unhide before delete — xlSheetVeryHidden sheets cannot be deleted directly
    On Error Resume Next
    ThisWorkbook.Worksheets(chartSheetName).Visible = xlSheetVisible
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(chartSheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsChart = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsChart.name = chartSheetName
    wsChart.Visible = xlSheetVisible
    wsChart.Activate
    DoEvents

    gStep = "BuildScratchPivotAndChart: Add chartobject " & chName
    Set coScratch = wsChart.ChartObjects.Add( _
        left:=wsChart.Cells(1, 1).left, _
        top:=wsChart.Cells(1, 1).top, _
        width:=750, height:=420)
    coScratch.name = chName

    Dim ch As Chart
    Set ch = coScratch.Chart

    ' Initialize with Budget+Actual via ChartWizard
    ' Source is on wsScratch (different sheet) — no pivot on wsChart = plain chart
    gStep = "BuildScratchPivotAndChart: ChartWizard"
    ch.ChartWizard _
        Source:=wsScratch.Cells(blkRow, blkCol).Resize(centerCount + 1, 3), _
        Gallery:=xlColumnClustered, _
        PlotBy:=xlColumns, _
        CategoryLabels:=1, _
        SeriesLabels:=1, _
        HasLegend:=True

    On Error Resume Next
    ch.SeriesCollection(1).name = "Total Budget"
    ch.SeriesCollection(2).name = "Total Actual"
    On Error GoTo EH

    ' NewSeries now works — chart is on plain sheet, not pivot sheet
    gStep = "BuildScratchPivotAndChart: add Ach% series via NewSeries"
    Dim sAch As Series
    Set sAch = ch.SeriesCollection.NewSeries
    sAch.name = "Ach %"
    sAch.Values = rAchPct
    sAch.XValues = rLabels

    Debug.Print "Series count: " & ch.SeriesCollection.Count

    gStep = "BuildScratchPivotAndChart: ApplyScratchChartFormatting"
    ApplyScratchChartFormatting ch

    ' Hide the chart temp sheet — it gets exported via AddChartSheetFromChart
    wsChart.Visible = xlSheetVeryHidden

    ' Hide scratch data sheet
    wsScratch.Visible = xlSheetVeryHidden

    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    On Error GoTo EH

    Exit Sub

EH:
    On Error Resume Next
    If Not wsChart Is Nothing Then wsChart.Visible = xlSheetVeryHidden
    wsScratch.Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("Dashboard").Activate
    On Error GoTo 0
    Err.Raise vbObjectError + 9200, , _
        "BuildScratchPivotAndChart [" & country & "] failed at: " & gStep & _
        " - " & Err.Number & ": " & Err.Description
End Sub


'--------------------------------------------------
' Export one page per region using the scratch pivot + scratch chart.
'
' @param exportNames Collection of CS__ sheets for final PDF
' @param country Country name/code
' @param ptScratch Scratch pivot
' @param chScratch Scratch chart (not a Shape)
' @param generatedAt Timestamp
'--------------------------------------------------
Private Sub ExportCountryByRegion_FromScratch( _
    ByVal exportNames As Collection, _
    ByVal country As String, _
    ByVal ptScratch As PivotTable, _
    ByVal chScratch As Chart, _
    ByVal generatedAt As Date)

    ' Region is now a PageField - get all its items
    Dim pfRegion As PivotField
    On Error Resume Next
    Set pfRegion = ptScratch.PivotFields("Region")
    On Error GoTo 0

    If pfRegion Is Nothing Then
        Debug.Print "Scratch pivot missing field [Region] for country [" & country & "]"
        Exit Sub
    End If

    ' Collect region names first (avoid modifying collection while iterating)
    Dim regionNames() As String
    Dim nRegions As Long
    Dim pi As PivotItem
    For Each pi In pfRegion.PivotItems
        If Not IsSkippablePivotItem(pi.name) Then
            nRegions = nRegions + 1
            ReDim Preserve regionNames(1 To nRegions)
            regionNames(nRegions) = pi.name
        End If
    Next pi

    If nRegions = 0 Then
        Debug.Print "No regions found for country [" & country & "]"
        Exit Sub
    End If

    ' Export one page per region
    Dim i As Long
    For i = 1 To nRegions
        gStep = "Scratch filter Region=" & regionNames(i) & " for " & country

        ' Set PageField to this region
        On Error Resume Next
        pfRegion.ClearAllFilters
        pfRegion.CurrentPage = regionNames(i)
        On Error GoTo 0

        ' Hide grand totals — we don't want them in the chart
        ptScratch.ColumnGrand = False
        ptScratch.RowGrand = False
    
        ptScratch.ManualUpdate = False
        ptScratch.RefreshTable
        DoEvents

        ' Rebuild Ach% static block after region filter change
        RebuildAchPctBlock ptScratch, chScratch

        ' Refresh chart
        On Error Resume Next
        chScratch.Refresh
        On Error GoTo 0
        DoEvents

        gStep = "Export scratch chart " & country & " - " & regionNames(i)
        AddChartSheetFromChart exportNames, _
            country & " - " & regionNames(i), _
            chScratch, generatedAt
    Next i

    ' Restore Region to (All)
    On Error Resume Next
    pfRegion.ClearAllFilters
    pfRegion.CurrentPage = "(All)"
    ptScratch.RefreshTable
    On Error GoTo 0
End Sub

Private Function IsSkippablePivotItem(ByVal itemName As String) As Boolean
    Dim nm As String
    nm = LCase$(Trim$(itemName))

    If nm = "(blank)" Then IsSkippablePivotItem = True: Exit Function
    If nm = "(all)" Then IsSkippablePivotItem = True: Exit Function
    If nm = "(multiple items)" Then IsSkippablePivotItem = True: Exit Function

    IsSkippablePivotItem = False
End Function
'--------------------------------------------------
' Add export sheet from a Chart object (not a Shape).
' Uses Chart.Export -> PNG (clipboard-free).
'
' @param names Collection of temp chart sheet names
' @param title Title in cell A1
' @param ch Chart object to export
' @param generatedAt Timestamp
'--------------------------------------------------
Private Sub AddChartSheetFromChart( _
    ByVal names As Collection, _
    ByVal title As String, _
    ByVal ch As Chart, _
    ByVal generatedAt As Date)

    gStep = "AddChartSheetFromChart: export chart to PNG"

    Dim tmpPng As String
    tmpPng = GetTempPngPath(CS_PREFIX & Format$(names.Count + 1, "000"))

    On Error Resume Next
    ch.Refresh
    On Error GoTo 0
    DoEvents

    If Not ch.Export(fileName:=tmpPng, FilterName:="PNG") Then
        Err.Raise vbObjectError + 2211, , "Chart.Export failed for [" & title & "]"
    End If

    gStep = "AddChartSheetFromChart: create temp worksheet"
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.name = CS_PREFIX & Format$(names.Count + 1, "000")

    wsOut.Cells.Clear
    wsOut.Range("A1").Value = title
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.size = 14

    gStep = "AddChartSheetFromChart: place PNG"
    Dim piC As Shape
    Set piC = wsOut.Shapes.AddPicture( _
        fileName:=tmpPng, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        left:=wsOut.Range("A3").left, _
        top:=wsOut.Range("A3").top, _
        width:=-1, _
        height:=-1)

    FitShapeToPage wsOut, piC, wsOut.Range("A3").left, wsOut.Range("A3").top

    With wsOut.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

    ApplyPdfFooter wsOut, generatedAt, "", ThisWorkbook.name

    Dim br As Range
    Set br = piC.BottomRightCell
    wsOut.PageSetup.PrintArea = wsOut.Range("A1", br).Address

    names.Add wsOut.name
End Sub

Private Sub CleanupScratchObjects(ByVal wsScratch As Worksheet, ByVal country As String)
    ' No-op now.
    ' Scratch sheet is reset at the start of each BuildScratchPivotAndChart.
End Sub
'--------------------------------------------------
' Clone Pivot layout WITHOUT causing repeated re-layout churn that can trigger:
' "1004 - A PivotTable report cannot overlap another PivotTable report."
'
' Strategy:
' 1) Freeze layout (ManualUpdate=True)
' 2) Hide all destination fields
' 3) Apply orientations (Page/Row/Column) only (NO Position yet)
' 4) Apply positions in a second pass
' 5) Refresh ONCE at the end
'
' @param ptFrom Source pivot
' @param ptTo   Destination (scratch) pivot
'--------------------------------------------------
Private Sub ClonePivotLayout(ByVal ptFrom As PivotTable, ByVal ptTo As PivotTable)
    On Error GoTo EH

    ptTo.ManualUpdate = True

    Dim pfFrom As PivotField
    For Each pfFrom In ptFrom.PivotFields

        ' Only copy meaningful pivot fields
        If pfFrom.Orientation = xlRowField _
           Or pfFrom.Orientation = xlColumnField _
           Or pfFrom.Orientation = xlPageField _
           Or pfFrom.Orientation = xlDataField Then

            Dim pfTo As PivotField
            Set pfTo = Nothing
            On Error Resume Next
            Set pfTo = ptTo.PivotFields(pfFrom.name)
            On Error GoTo EH

            If Not pfTo Is Nothing Then
                pfTo.Orientation = pfFrom.Orientation
                If pfFrom.Orientation <> xlDataField Then
                    pfTo.Position = pfFrom.Position
                End If
            End If
        End If
    Next pfFrom

    ptTo.ManualUpdate = False
    ptTo.RefreshTable
    Exit Sub

EH:
    ' Bubble with context
    Err.Raise Err.Number, , "ClonePivotLayout failed: " & Err.Description
End Sub
'--------------------------------------------------
' Set a field's Orientation if the field exists in the destination pivot.
' @param pt Destination pivot
' @param fieldName Field name
' @param orient One of xlPageField/xlRowField/xlColumnField/xlDataField/xlHidden
'--------------------------------------------------
Private Sub SetPivotFieldOrientation_IfExists(ByVal pt As PivotTable, ByVal fieldName As String, ByVal orient As XlPivotFieldOrientation)
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0
    If pf Is Nothing Then Exit Sub

    ' IMPORTANT: we do NOT handle DataField cloning here (keep it simple & safe)
    If orient = xlDataField Then Exit Sub

    On Error Resume Next
    pf.Orientation = orient
    On Error GoTo 0
End Sub

'--------------------------------------------------
' Set a field's Position if the field exists in the destination pivot.
' @param pt Destination pivot
' @param fieldName Field name
' @param pos Position (1..n)
'--------------------------------------------------
Private Sub SetPivotFieldPosition_IfExists(ByVal pt As PivotTable, ByVal fieldName As String, ByVal pos As Long)
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0
    If pf Is Nothing Then Exit Sub

    On Error Resume Next
    pf.Position = pos
    On Error GoTo 0
End Sub

Private Sub ResetScratchSheet(ByVal ws As Worksheet)
    gStep = "ResetScratchSheet: " & ws.name

    On Error Resume Next

    ' Must be visible to reliably delete objects
    ws.Visible = xlSheetVisible

    Application.DisplayAlerts = False

    ' Delete chart objects
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    ' Delete shapes (pictures, slicers, etc.)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Nuke all cells (this also removes pivot table reports)
    ws.Cells.Clear
    ws.Cells.ClearFormats

    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
'--------------------------------------------------
' Hard reset scratch sheet (delete + recreate) to ensure
' ZERO pivot footprints remain (prevents overlap error).
'
' @return Fresh scratch worksheet
'--------------------------------------------------
Private Function RecreateScratchPivotSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET)
    On Error GoTo 0

    Application.DisplayAlerts = False
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        ws.Delete
    End If
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.name = SCRATCH_SHEET
    ws.Cells.Clear

    ' Leave visible — caller controls visibility
    ws.Visible = xlSheetVisible

    Set RecreateScratchPivotSheet = ws
End Function
'--------------------------------------------------
' Delete + recreate the scratch sheet.
' This is the only reliable way to avoid "overlap" caused by
' lingering PivotTable objects that survive Cells.Clear.
'
' @return A brand new worksheet instance named SCRATCH_SHEET
'--------------------------------------------------
Private Function CreateFreshScratchPivotSheet() As Worksheet
    Dim ws As Worksheet
    Dim attempt As Long

    ' Step 1: force-delete existing scratch sheet, handling structure protection
    Application.DisplayAlerts = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Unprotect workbook structure if needed (sheet.Delete requires this)
        Dim wasStructProtected As Boolean
        wasStructProtected = ThisWorkbook.ProtectStructure
        If wasStructProtected Then
            On Error Resume Next
            ThisWorkbook.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
            On Error GoTo 0
        End If

        ws.Visible = xlSheetVisible

        ' Delete all PivotTables on the sheet FIRST (belt-and-suspenders)
        Dim pt As PivotTable
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.TableRange2.Delete
            On Error GoTo 0
        Next pt

        On Error Resume Next
        ws.Delete
        On Error GoTo 0

        ' Re-protect structure if it was protected before
        If wasStructProtected Then
            On Error Resume Next
            ThisWorkbook.Protect Password:=modSecurity.SYS_STRUCT_PWD, Structure:=True
            On Error GoTo 0
        End If
    End If
    Application.DisplayAlerts = True

    ' Step 2: verify it is truly gone before adding fresh one
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Sheet still exists - delete all its pivots and clear it instead
        Application.DisplayAlerts = False
        Dim pt2 As PivotTable
        For Each pt2 In ws.PivotTables
            On Error Resume Next
            pt2.TableRange2.Delete
            On Error GoTo 0
        Next pt2
        ws.Cells.Clear
        ws.Cells.ClearFormats
        Application.DisplayAlerts = True
    Else
        ' Step 3: add a fresh sheet
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SCRATCH_SHEET
    End If

    ws.Visible = xlSheetVeryHidden
    Set CreateFreshScratchPivotSheet = ws
End Function
'========================================================
' Debug helper: Region field orientation
'========================================================
Public Sub DebugPrintRegionOrientation(ByVal pt As PivotTable, ByVal label As String)
    On Error GoTo EH

    Dim pf As PivotField
    Set pf = pt.PivotFields("Region")

    Debug.Print "=== " & label & " ==="
    Debug.Print "Pivot: " & pt.name & "  Sheet: " & pt.Parent.name
    Debug.Print "Region Orientation: " & PivotOrientationName(pf.Orientation) & "  Value=" & CStr(pf.Orientation)
    Debug.Print "EnableMultiplePageItems: " & CStr(pf.EnableMultiplePageItems)
    Debug.Print "PivotItems.Count: " & CStr(pf.PivotItems.Count)
    Debug.Print "----------------------"
    Exit Sub

EH:
    Debug.Print "=== " & label & " ==="
    Debug.Print "Pivot: " & pt.name & "  Sheet: " & pt.Parent.name
    Debug.Print "Region field not found or error: " & Err.Number & " - " & Err.Description
    Debug.Print "----------------------"
End Sub

Private Function PivotOrientationName(ByVal o As XlPivotFieldOrientation) As String
    Select Case o
        Case xlRowField: PivotOrientationName = "xlRowField"
        Case xlColumnField: PivotOrientationName = "xlColumnField"
        Case xlPageField: PivotOrientationName = "xlPageField"
        Case xlDataField: PivotOrientationName = "xlDataField"
        Case xlHidden: PivotOrientationName = "xlHidden"
        Case Else: PivotOrientationName = "Unknown(" & CStr(o) & ")"
    End Select
End Function

Public Sub DebugPrintPivotFieldPlacements(ByVal label As String, ByVal pt As PivotTable)
    Debug.Print "=== " & label & " placements (" & pt.name & ") ==="

    Dim pf As PivotField

    Debug.Print "RowFields:"
    For Each pf In pt.RowFields
        Debug.Print "  - " & pf.name & " (pos " & pf.Position & ")"
    Next pf

    Debug.Print "ColumnFields:"
    For Each pf In pt.ColumnFields
        Debug.Print "  - " & pf.name & " (pos " & pf.Position & ")"
    Next pf

    Debug.Print "PageFields:"
    For Each pf In pt.PageFields
        Debug.Print "  - " & pf.name & " (pos " & pf.Position & ")"
    Next pf

    Debug.Print "DataFields:"
    For Each pf In pt.DataFields
        Debug.Print "  - " & pf.caption & " / Source=" & pf.SourceName
    Next pf

    Debug.Print "----------------------"
End Sub

'========================================================
' Scratch sheet: ALWAYS recreate to guarantee no ghost pivots / overlap footprints
'========================================================
Private Function GetOrCreateScratchPivotSheet() As Worksheet
    Dim ws As Worksheet

    gStep = "GetOrCreateScratchPivotSheet: recreate"

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SCRATCH_SHEET).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(SCRATCH_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.name = SCRATCH_SHEET
    ws.Cells.Clear
    ws.Cells.ClearFormats

    ' Keep it out of UI
    ws.Visible = xlSheetVeryHidden

    Set GetOrCreateScratchPivotSheet = ws
End Function
Private Sub ResetScratchPivotSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
Private Sub ClonePivotLayout_Safe_AddFields(ByVal ptFrom As PivotTable, ByVal ptTo As PivotTable, _
                                             Optional ByVal countryFilter As String = "")
    On Error GoTo EH

    ptTo.ManualUpdate = True

    Dim arrRow() As String, arrCol() As String, arrPage() As String
    Dim nRow As Long, nCol As Long, nPage As Long
    Dim pf As PivotField

    ' Collect Row fields - EXCLUDE Country (force it as PageField for filtering)
    For Each pf In ptFrom.RowFields
        If pf.name <> "Country" Then
            nRow = nRow + 1
            ReDim Preserve arrRow(1 To nRow)
            arrRow(nRow) = pf.name
        End If
    Next pf

    For Each pf In ptFrom.ColumnFields
        If pf.name <> "Country" Then
            nCol = nCol + 1
            ReDim Preserve arrCol(1 To nCol)
            arrCol(nCol) = pf.name
        End If
    Next pf

    For Each pf In ptFrom.PageFields
        If pf.name <> "Country" Then
            nPage = nPage + 1
            ReDim Preserve arrPage(1 To nPage)
            arrPage(nPage) = pf.name
        End If
    Next pf

    ' Always add Country as PageField so CurrentPage filter works
    nPage = nPage + 1
    ReDim Preserve arrPage(1 To nPage)
    arrPage(nPage) = "Country"

    Dim vRow As Variant, vCol As Variant, vPage As Variant
    If nRow > 0 Then vRow = arrRow Else vRow = Empty
    If nCol > 0 Then vCol = arrCol Else vCol = Empty
    If nPage > 0 Then vPage = arrPage Else vPage = Empty

    gStep = "ClonePivotLayout: AddFields"
    ptTo.AddFields RowFields:=vRow, ColumnFields:=vCol, PageFields:=vPage

    ' Add Budget and Actual data fields only (skip AchievementPct - we compute Ach% ourselves)
    Dim df As PivotField
    For Each df In ptFrom.DataFields
        If InStr(1, df.SourceName, "Achievement", vbTextCompare) = 0 And _
           InStr(1, df.SourceName, "Ach", vbTextCompare) = 0 Then
            gStep = "ClonePivotLayout: AddDataField " & df.SourceName
            Dim added As PivotField
            Set added = ptTo.AddDataField(ptTo.PivotFields(df.SourceName), df.caption, df.Function)
            On Error Resume Next
            added.NumberFormat = df.NumberFormat
            On Error GoTo EH
        End If
    Next df

    ' MUST refresh with ManualUpdate=False BEFORE adding CalculatedField
    ptTo.ManualUpdate = False
    ptTo.RefreshTable
    DoEvents

    ' Add Ach % as calculated field = Actual / Budget
    gStep = "ClonePivotLayout: Add calculated Ach %"

    ' Safe delete if already exists
    Dim cf As Object
    On Error Resume Next
    For Each cf In ptTo.CalculatedFields
        If cf.name = "Ach %" Then
            cf.Delete
            Exit For
        End If
    Next cf
    On Error GoTo EH

    ' Add the calculated field
    ptTo.CalculatedFields.Add name:="Ach %", Formula:="=Actual/Budget"
    ptTo.RefreshTable
    DoEvents

    ' Surface as data field - find it by looping PivotFields since name lookup is unreliable
    gStep = "ClonePivotLayout: AddDataField Ach %"
    Dim addedCalc As PivotField
    Dim pfSearch As PivotField
    On Error Resume Next
    For Each pfSearch In ptTo.PivotFields
        If pfSearch.name = "Ach %" Then
            Set addedCalc = ptTo.AddDataField(pfSearch, "Ach %", xlSum)
            Exit For
        End If
    Next pfSearch
    On Error GoTo EH

    If addedCalc Is Nothing Then
        ' Fallback: try direct name access with Sum prefix Excel sometimes applies
        On Error Resume Next
        Set addedCalc = ptTo.AddDataField(ptTo.PivotFields("Ach %"), "Ach %", xlSum)
        If addedCalc Is Nothing Then
            Set addedCalc = ptTo.AddDataField(ptTo.PivotFields("Sum of Ach %"), "Ach %", xlSum)
        End If
        On Error GoTo EH
    End If

    If Not addedCalc Is Nothing Then
        addedCalc.NumberFormat = "0.0%"
    End If

    ptTo.RefreshTable
    DoEvents
    
    ' Apply country page filter using PivotItems (more reliable than CurrentPage)
    If Len(countryFilter) > 0 Then
        gStep = "ClonePivotLayout: apply country filter=" & countryFilter

        ptTo.ManualUpdate = True

        Dim pfCountry As PivotField
        Set pfCountry = ptTo.PivotFields("Country")

        ' First try CurrentPage (works when field is already a clean PageField)
        On Error Resume Next
        pfCountry.ClearAllFilters
        pfCountry.CurrentPage = countryFilter
        Dim currentPageErr As Long
        currentPageErr = Err.Number
        On Error GoTo EH

        ' If CurrentPage failed, fall back to manual PivotItem visibility
        If currentPageErr <> 0 Then
            gStep = "ClonePivotLayout: apply country filter via PivotItems=" & countryFilter
            pfCountry.EnableMultiplePageItems = True

            Dim pi As PivotItem
            On Error Resume Next
            For Each pi In pfCountry.PivotItems
                If StrComp(Trim$(pi.name), Trim$(countryFilter), vbTextCompare) = 0 Then
                    pi.Visible = True
                Else
                    pi.Visible = False
                End If
            Next pi
            On Error GoTo EH
        End If

        ptTo.ManualUpdate = False
        ptTo.RefreshTable
        DoEvents
    End If

    ptTo.RefreshTable
    Exit Sub

EH:
    On Error Resume Next
    ptTo.ManualUpdate = False
    ptTo.RefreshTable
    On Error GoTo 0
    Err.Raise vbObjectError + 9123, , "ClonePivotLayout failed at [" & gStep & "]: " & Err.Number & " - " & Err.Description
End Sub
'--------------------------------------------------
' Hide all PivotChart field buttons and display labels
' (removes Month/Country dropdowns, field headers,
'  axis labels, and expand/collapse +/- buttons)
'--------------------------------------------------
Private Sub HidePivotChartButtons(ByVal ch As Chart)
    On Error Resume Next

    ' Hide all pivot field buttons (the dropdown buttons on the chart)
    Dim btn As PivotLayout
    Set btn = Nothing

    ' Turn off every display button type on the chart
    ch.ShowReportFilterFieldButtons = False
    ch.ShowLegendFieldButtons = False
    ch.ShowAxisFieldButtons = False
    ch.ShowValueFieldButtons = False
    ch.ShowExpandCollapseEntireFieldButtons = False

    On Error GoTo 0
End Sub
'--------------------------------------------------
' Copy AxisGroup (primary/secondary) from each series
' in the original PivotChart to the scratch chart.
' Matches by series index position.
'--------------------------------------------------
Private Sub CloneSeriesAxisGroups(ByVal chFrom As Chart, ByVal chTo As Chart)
    On Error Resume Next

    Dim i As Long
    Dim maxSeries As Long
    maxSeries = chFrom.SeriesCollection.Count
    If chTo.SeriesCollection.Count < maxSeries Then
        maxSeries = chTo.SeriesCollection.Count
    End If

    For i = 1 To maxSeries
        chTo.SeriesCollection(i).AxisGroup = chFrom.SeriesCollection(i).AxisGroup
    Next i

    On Error GoTo 0
End Sub
'--------------------------------------------------
' Find the "Ach %" series in a chart and style it:
' - Bar/column fill: gray (RGB 128,128,128)
' - Legend entry label: white text on gray background
' - Applies to both the series bars AND the legend key
'--------------------------------------------------
Private Sub StyleAchPctSeries(ByVal ch As Chart)
    On Error Resume Next

    Dim sr As Series
    For Each sr In ch.SeriesCollection
        If InStr(1, sr.name, "Ach", vbTextCompare) > 0 Then

            ' Gray fill on the series bars
            sr.Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
            sr.Format.Fill.BackColor.RGB = RGB(128, 128, 128)
            sr.Format.Fill.Solid

            ' Style data labels if present
            On Error Resume Next
            Dim dl As DataLabel
            For Each dl In sr.DataLabels
                dl.Font.Color = RGB(255, 255, 255)
                dl.Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
                dl.Format.Fill.Solid
            Next dl
            On Error GoTo 0

        End If
    Next sr

    ' Style the legend entry for Ach %
    Dim le As LegendEntry
    Dim li As LegendKey
    Dim lg As Legend
    Set lg = ch.Legend

    Dim idx As Long
    For idx = 1 To ch.SeriesCollection.Count
        If InStr(1, ch.SeriesCollection(idx).name, "Ach", vbTextCompare) > 0 Then
            Set le = lg.LegendEntries(idx)
            ' White text
            le.Font.Color = RGB(255, 255, 255)
            ' Gray background on legend entry
            le.Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
            le.Format.Fill.Solid
            Exit For
        End If
    Next idx

    On Error GoTo 0
End Sub
'--------------------------------------------------
' Apply full formatting to scratch chart to match
' the original dashboard PivotChart style:
'   Series 1 = Total Budget  -> clustered column, primary axis, blue
'   Series 2 = Total Actual  -> clustered column, primary axis, orange
'   Series 3 = Ach %         -> line, SECONDARY axis, gray with data labels
'                               gray box label background, white text
'--------------------------------------------------
Private Sub ApplyScratchChartFormatting(ByVal ch As Chart)
    On Error Resume Next

    ' Debug: print series count and names to Immediate window
    Debug.Print "ApplyScratchChartFormatting: SeriesCollection.Count = " & ch.SeriesCollection.Count
    Dim dbg As Long
    For dbg = 1 To ch.SeriesCollection.Count
        Debug.Print "  Series(" & dbg & ") = " & ch.SeriesCollection(dbg).name
    Next dbg

    ch.chartType = xlColumnClustered

    Dim sc As SeriesCollection
    Set sc = ch.SeriesCollection

    ' ===== Series 1: Total Budget =====
    If sc.Count >= 1 Then
        With sc(1)
            .chartType = xlColumnClustered
            .AxisGroup = xlPrimary
            .Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
            .Format.Fill.Solid
            .Format.Line.Visible = msoFalse
            .HasDataLabels = True
            With .DataLabels
                .ShowValue = True
                .Position = xlLabelPositionOutsideEnd
                .Font.size = 7
                .Font.Color = RGB(0, 0, 0)
                .Format.Fill.Visible = msoFalse
                .NumberFormat = "#,##0"
            End With
        End With
    End If

    ' ===== Series 2: Total Actual =====
    If sc.Count >= 2 Then
        With sc(2)
            .chartType = xlColumnClustered
            .AxisGroup = xlPrimary
            .Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
            .Format.Fill.Solid
            .Format.Line.Visible = msoFalse
            .HasDataLabels = True
            With .DataLabels
                .ShowValue = True
                .Position = xlLabelPositionOutsideEnd
                .Font.size = 7
                .Font.Color = RGB(0, 0, 0)
                .Format.Fill.Visible = msoFalse
                .NumberFormat = "#,##0"
            End With
        End With
    End If

    ' ===== Series 3: Ach % — secondary axis LINE =====
    If sc.Count >= 3 Then
        Debug.Print "  Setting Series 3 to secondary axis line"
        With sc(3)
            ' AxisGroup MUST be set before ChartType in Excel 2019
            .AxisGroup = xlSecondary
            .chartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(128, 128, 128)
            .Format.Line.Weight = 1.5
            .MarkerStyle = -4142   ' xlMarkerStyleNone
            .HasDataLabels = True
            With .DataLabels
                .ShowValue = True
                .NumberFormat = "0.0%"
                .Font.size = 7
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Format.Fill.ForeColor.RGB = RGB(89, 89, 89)
                .Format.Fill.Solid
                .Format.Line.Visible = msoFalse
            End With
        End With
    Else
        Debug.Print "  WARNING: sc.Count=" & sc.Count & " - Ach% series NOT found!"
    End If

    ' ===== Primary Y axis =====
    With ch.Axes(xlValue, xlPrimary)
        .HasTitle = False
        .TickLabels.NumberFormat = "#,##0"
        .TickLabels.Font.size = 8
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
    End With

    ' ===== Secondary Y axis (% scale) =====
    If sc.Count >= 3 Then
        With ch.Axes(xlValue, xlSecondary)
            .HasTitle = False
            .TickLabels.NumberFormat = "0%"
            .TickLabels.Font.size = 8
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        End With
    End If

    ' ===== Category X axis =====
    With ch.Axes(xlCategory)
        .TickLabels.Font.size = 8
        .TickLabels.Font.Color = RGB(0, 0, 0)
    End With

    ' ===== Plot area border =====
    With ch.PlotArea.Border
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' ===== Legend at bottom =====
    With ch.Legend
        .Position = xlLegendPositionBottom
        .Font.size = 8
    End With

    ' ===== Hide all pivot field buttons =====
    ch.ShowReportFilterFieldButtons = False
    ch.ShowLegendFieldButtons = False
    ch.ShowAxisFieldButtons = False
    ch.ShowValueFieldButtons = False
    ch.ShowExpandCollapseEntireFieldButtons = False

    On Error GoTo 0
End Sub
'--------------------------------------------------
' Rebuilds the static chart data block (col F-I on wsScratch)
' after pivot region filter changes.
' Called once per region before chart export.
'--------------------------------------------------
Private Sub RebuildAchPctBlock(ByVal ptScratch As PivotTable, ByVal chScratch As Chart)
    On Error Resume Next

    Dim pvtBody As Range
    Set pvtBody = ptScratch.DataBodyRange
    If pvtBody Is Nothing Then Exit Sub
    
    ' Ensure grand totals off so DataBodyRange excludes grand total row
    On Error Resume Next
    ptScratch.RowGrand = False
    ptScratch.ColumnGrand = False
    On Error GoTo 0
    
    Dim centerCount As Long
    centerCount = pvtBody.Rows.Count

    Dim wsScratch As Worksheet
    Set wsScratch = ptScratch.Parent  ' pivot's parent sheet

    Dim blkCol As Long: blkCol = 6   ' col F — must match BuildScratchPivotAndChart
    Dim blkRow As Long: blkRow = 1

    ' Clear old rows first so previous-region values cannot leak into chart.
    wsScratch.Range(wsScratch.Cells(blkRow + 1, blkCol), wsScratch.Cells(blkRow + 5000, blkCol + 3)).ClearContents

    ' Keep headers stable.
    wsScratch.Cells(blkRow, blkCol).Value = "Center"
    wsScratch.Cells(blkRow, blkCol + 1).Value = "Total Budget"
    wsScratch.Cells(blkRow, blkCol + 2).Value = "Total Actual"
    wsScratch.Cells(blkRow, blkCol + 3).Value = "Ach %"


    Dim rowLabelRange As Range
    Set rowLabelRange = ptScratch.rowRange

    Dim i As Long
    Dim bv As Double, av As Double
    For i = 1 To centerCount
        bv = 0: av = 0
        bv = CDbl(pvtBody.Cells(i, 1).Value)
        av = CDbl(pvtBody.Cells(i, 2).Value)
        wsScratch.Cells(blkRow + i, blkCol).Value = _
            CStr(rowLabelRange.Cells(i + 1, 1).Value)
        wsScratch.Cells(blkRow + i, blkCol + 1).Value = bv
        wsScratch.Cells(blkRow + i, blkCol + 2).Value = av
        If bv <> 0 Then
            wsScratch.Cells(blkRow + i, blkCol + 3).Value = av / bv
        Else
            wsScratch.Cells(blkRow + i, blkCol + 3).Value = 0
        End If
    Next i

    ' Rebind chart series to this exact region-sized block.
    Dim rLabels As Range, rBudget As Range, rActual As Range, rAchPct As Range
    Set rLabels = wsScratch.Cells(blkRow + 1, blkCol).Resize(centerCount, 1)
    Set rBudget = wsScratch.Cells(blkRow + 1, blkCol + 1).Resize(centerCount, 1)
    Set rActual = wsScratch.Cells(blkRow + 1, blkCol + 2).Resize(centerCount, 1)
    Set rAchPct = wsScratch.Cells(blkRow + 1, blkCol + 3).Resize(centerCount, 1)

    If Not chScratch Is Nothing Then
        If chScratch.SeriesCollection.Count >= 1 Then
            chScratch.SeriesCollection(1).name = "Total Budget"
            chScratch.SeriesCollection(1).XValues = rLabels
            chScratch.SeriesCollection(1).Values = rBudget
        End If
        If chScratch.SeriesCollection.Count >= 2 Then
            chScratch.SeriesCollection(2).name = "Total Actual"
            chScratch.SeriesCollection(2).XValues = rLabels
            chScratch.SeriesCollection(2).Values = rActual
        End If
        If chScratch.SeriesCollection.Count >= 3 Then
            chScratch.SeriesCollection(3).name = "Ach %"
            chScratch.SeriesCollection(3).XValues = rLabels
            chScratch.SeriesCollection(3).Values = rAchPct
        End If
    End If

    On Error GoTo 0
End Sub



'--------------------------------------------------
' Adds a "Monthly Budget vs Actual" chart page for
' a specific country. Sets ptMonth Country PageField
' to the country, reads data row by row, resets filter.
'--------------------------------------------------
Private Sub AddCountryMonthlyChartPage( _
    ByVal names As Collection, _
    ByVal country As String, _
    ByVal generatedAt As Date)

    On Error GoTo EH

    gStep = "AddCountryMonthlyChartPage [" & country & "]: aggregate month data from tblData"

    Dim loSrc As ListObject
    On Error Resume Next
    Set loSrc = ThisWorkbook.Worksheets("Data").ListObjects("tblData")
    On Error GoTo EH
    If loSrc Is Nothing Then
        Debug.Print "AddCountryMonthlyChartPage: tblData not found"
        Exit Sub
    End If

    Dim srcBody As Range
    Set srcBody = loSrc.DataBodyRange
    If srcBody Is Nothing Then
        Debug.Print "AddCountryMonthlyChartPage: tblData has no rows"
        Exit Sub
    End If

    Dim colCountry As Long, colMonth As Long, colBudget As Long, colActual As Long
    Dim hdr As ListColumn
    For Each hdr In loSrc.ListColumns
        Select Case LCase$(hdr.name)
            Case "country": colCountry = hdr.Index
            Case "month":   colMonth = hdr.Index
            Case "budget":  colBudget = hdr.Index
            Case "actual":  colActual = hdr.Index
        End Select
    Next hdr
    If colCountry = 0 Or colMonth = 0 Or colBudget = 0 Or colActual = 0 Then
        Debug.Print "AddCountryMonthlyChartPage: required columns missing for " & country
        Exit Sub
    End If

    Dim monthMap As Object
    Set monthMap = CreateObject("Scripting.Dictionary")
    monthMap.CompareMode = vbTextCompare

    Dim countryKey As String
    countryKey = UCase$(Trim$(country))

    Dim r As Long
    Dim bVal As Double, aVal As Double
    For r = 1 To srcBody.Rows.Count
        If UCase$(Trim$(CStr(srcBody.Cells(r, colCountry).Value2))) = countryKey Then
            Dim monthRaw As Variant
            monthRaw = srcBody.Cells(r, colMonth).Value2
            If Len(Trim$(CStr(monthRaw))) = 0 Then GoTo NextSrcRow

            bVal = 0: aVal = 0
            On Error Resume Next
            bVal = CDbl(srcBody.Cells(r, colBudget).Value2)
            aVal = CDbl(srcBody.Cells(r, colActual).Value2)
            Err.Clear
            On Error GoTo EH

            Dim monthDisp As Variant
            Dim mapKey As String
            Dim sortKey As String
            If IsDate(monthRaw) Then
                monthDisp = CDate(monthRaw)
                sortKey = "D|" & Format$(CDate(monthDisp), "yyyymmdd")
                mapKey = sortKey
            ElseIf IsNumeric(monthRaw) Then
                monthDisp = DateAdd("d", CLng(CDbl(monthRaw)), #12:00:00 AM#)
                sortKey = "D|" & Format$(CDate(monthDisp), "yyyymmdd")
                mapKey = sortKey
            Else
                monthDisp = Trim$(CStr(monthRaw))
                sortKey = "T|" & UCase$(CStr(monthDisp))
                mapKey = sortKey
            End If

            Dim rec As Variant
            If monthMap.Exists(mapKey) Then
                rec = monthMap(mapKey)
                rec(1) = CDbl(rec(1)) + bVal
                rec(2) = CDbl(rec(2)) + aVal
                monthMap(mapKey) = rec
            Else
                Dim newRec(0 To 3) As Variant
                newRec(0) = monthDisp
                newRec(1) = bVal
                newRec(2) = aVal
                newRec(3) = sortKey
                monthMap.Add mapKey, newRec
            End If
        End If
NextSrcRow:
    Next r

    If monthMap.Count = 0 Then
        Debug.Print "AddCountryMonthlyChartPage: no data rows for " & country
        Exit Sub
    End If

    Dim monthKeys() As String
    ReDim monthKeys(1 To monthMap.Count)
    Dim k As Variant
    Dim idx As Long
    idx = 1
    For Each k In monthMap.keys
        monthKeys(idx) = CStr(k)
        idx = idx + 1
    Next k
    SortMonthAggregateKeys monthKeys, monthMap

    ' Use a dedicated temp sheet for chart data - avoids PivotTable row conflicts
    Dim wsScratch As Worksheet
    Dim tmpSheetName As String
    tmpSheetName = "CS__MonthlyTmp"
    On Error Resume Next
    Set wsScratch = ThisWorkbook.Worksheets(tmpSheetName)
    On Error GoTo EH
    If wsScratch Is Nothing Then
        Set wsScratch = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsScratch.name = tmpSheetName
    End If
    wsScratch.Cells.Clear

    Dim pasteCell As Range
    Set pasteCell = wsScratch.Range("A1")
    pasteCell.CurrentRegion.Clear

    ' Header row
    pasteCell.Offset(0, 0).Value = "Month"
    pasteCell.Offset(0, 1).Value = "Total Budget"
    pasteCell.Offset(0, 2).Value = "Total Actual"
    pasteCell.Offset(0, 3).Value = "Ach %"

    Dim iRow As Long
    iRow = 1
    For idx = LBound(monthKeys) To UBound(monthKeys)
        Dim rowRec As Variant
        rowRec = monthMap(monthKeys(idx))

        iRow = iRow + 1
        If IsDate(rowRec(0)) Then
            pasteCell.Offset(iRow - 1, 0).Value = CDate(rowRec(0))
            pasteCell.Offset(iRow - 1, 0).NumberFormat = "mmm-yy"
        ElseIf IsNumeric(rowRec(0)) Then
            pasteCell.Offset(iRow - 1, 0).Value = DateAdd("d", CLng(CDbl(rowRec(0))), #12:00:00 AM#)
            pasteCell.Offset(iRow - 1, 0).NumberFormat = "mmm-yy"
        Else
            pasteCell.Offset(iRow - 1, 0).Value = CStr(rowRec(0))
        End If

        bVal = CDbl(rowRec(1))
        aVal = CDbl(rowRec(2))
        pasteCell.Offset(iRow - 1, 1).Value = bVal
        pasteCell.Offset(iRow - 1, 1).NumberFormat = "#,##0"
        pasteCell.Offset(iRow - 1, 2).Value = aVal
        pasteCell.Offset(iRow - 1, 2).NumberFormat = "#,##0"
        pasteCell.Offset(iRow - 1, 3).Value = IIf(bVal <> 0, aVal / bVal, 0)
        pasteCell.Offset(iRow - 1, 3).NumberFormat = "0.0%"
    Next idx

    Dim dataRng As Range
    Set dataRng = pasteCell.Resize(iRow, 4)

    ' --- Build chart on scratch sheet ---
    gStep = "AddCountryMonthlyChartPage [" & country & "]: build chart"
    Dim chName As String
    chName = SCR_CH_PREFIX & "Monthly_" & SanitizeName_Local(country)

    Dim co As ChartObject
    On Error Resume Next
    Set co = wsScratch.ChartObjects(chName)
    If Not co Is Nothing Then co.Delete
    Set co = Nothing
    On Error GoTo EH

    Set co = wsScratch.ChartObjects.Add(left:=10, top:=10, width:=650, height:=360)
    co.name = chName

    ' Activate the chart before SetSourceData - required or it fails
    co.Activate
    DoEvents

    Dim ch As Chart
    Set ch = co.Chart

    ' Clear any auto-assigned series first
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim rngMonths As Range
    Dim rngBudget As Range
    Dim rngActual As Range
    Dim rngAch As Range
    Set rngMonths = pasteCell.Offset(1, 0).Resize(iRow - 1, 1)
    Set rngBudget = pasteCell.Offset(1, 1).Resize(iRow - 1, 1)
    Set rngActual = pasteCell.Offset(1, 2).Resize(iRow - 1, 1)
    Set rngAch = pasteCell.Offset(1, 3).Resize(iRow - 1, 1)
    Dim arrMonthLabels() As String
    Dim m As Long
    ReDim arrMonthLabels(1 To iRow - 1)
    For m = 1 To iRow - 1
        If IsDate(rngMonths.Cells(m, 1).Value) Then
            arrMonthLabels(m) = Format$(CDate(rngMonths.Cells(m, 1).Value), "mmm-yy")
        ElseIf IsNumeric(rngMonths.Cells(m, 1).Value) Then
            arrMonthLabels(m) = Format$(DateAdd("d", CLng(CDbl(rngMonths.Cells(m, 1).Value)), #12:00:00 AM#), "mmm-yy")
        Else
            arrMonthLabels(m) = CStr(rngMonths.Cells(m, 1).Value)
        End If
    Next m

    ch.chartType = xlColumnClustered

    Dim sBudget As Series
    Dim sActual As Series
    Dim sAch As Series
    Set sBudget = ch.SeriesCollection.NewSeries
    Set sActual = ch.SeriesCollection.NewSeries
    Set sAch = ch.SeriesCollection.NewSeries

    ' Series 1: Total Budget - primary column
    With sBudget
        .name = "Total Budget"
        .XValues = arrMonthLabels
        .Values = rngBudget
        .AxisGroup = xlPrimary
        .chartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Format.Fill.Solid
        .HasDataLabels = True
        .DataLabels.ShowValue = True
        .DataLabels.NumberFormat = "#,##0"
        .DataLabels.Font.size = 8
        .DataLabels.Format.Fill.Visible = msoFalse
    End With

    ' Series 2: Total Actual - primary column
    With sActual
        .name = "Total Actual"
        .XValues = arrMonthLabels
        .Values = rngActual
        .AxisGroup = xlPrimary
        .chartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .Format.Fill.Solid
        .HasDataLabels = True
        .DataLabels.ShowValue = True
        .DataLabels.NumberFormat = "#,##0"
        .DataLabels.Font.size = 8
        .DataLabels.Format.Fill.Visible = msoFalse
    End With

    ' Series 3: Ach% - secondary line
    sAch.AxisGroup = xlSecondary
    ch.HasAxis(xlValue, xlSecondary) = True
    With sAch
        .name = "Ach %"
        .XValues = arrMonthLabels
        .Values = rngAch
        .chartType = xlLine
        .AxisGroup = xlSecondary
        .Format.Line.ForeColor.RGB = RGB(128, 128, 128)
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
        .HasDataLabels = True
        .DataLabels.ShowValue = True
        .DataLabels.NumberFormat = "0.0%"
        .DataLabels.Font.size = 8
        .DataLabels.Font.Bold = True
        .DataLabels.Font.Color = RGB(255, 255, 255)
        .DataLabels.Format.Fill.ForeColor.RGB = RGB(89, 89, 89)
        .DataLabels.Format.Fill.Solid
        .DataLabels.Format.Line.Visible = msoFalse
    End With
    ' Axes
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
    ch.Axes(xlValue, xlSecondary).HasTitle = False
    ch.Axes(xlCategory).TickLabels.NumberFormatLinked = False
    ch.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"

    ' Title + legend
    ch.HasTitle = True
    ch.chartTitle.Text = "Monthly Budget vs Actual - " & country
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.size = 8

    ' --- Export PNG ---
    gStep = "AddCountryMonthlyChartPage [" & country & "]: export PNG"
    Dim tmpPng As String
    tmpPng = GetTempPngPath(CS_PREFIX & "mth_" & SanitizeName_Local(country))

    On Error Resume Next
    ch.Refresh
    On Error GoTo EH
    DoEvents

    If Not ch.Export(fileName:=tmpPng, FilterName:="PNG") Then
        Err.Raise vbObjectError + 2205, , "Chart.Export failed for [" & country & "]"
    End If

    ' --- Build export worksheet page ---
    gStep = "AddCountryMonthlyChartPage [" & country & "]: create page"
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.name = CS_PREFIX & Format$(names.Count + 1, "000")
    wsOut.Cells.Clear

    ' Title only - no inline date (moved to footer)
    wsOut.Range("A1").Value = country & " - Monthly Budget vs Actual (Overall)"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.size = 14

    ' Page setup FIRST so margins are set before placing picture
    With wsOut.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        ' Footer: generated date center, page # right
        .CenterFooter = "Generated: " & Format(generatedAt, "dd-mmm-yyyy hh:mm")
        .RightFooter = "Page &P of &N"
        .LeftFooter = ""
    End With

    ' Place chart PNG - use page-width size, FitToPagesWide=1 handles scaling
    Dim piC As Shape
    Set piC = wsOut.Shapes.AddPicture( _
        fileName:=tmpPng, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        left:=wsOut.Columns("A").left, _
        top:=wsOut.Rows(3).top, _
        width:=650, height:=360)

    ' Set print area to cover title + chart
    Dim lastRow As Long
    lastRow = wsOut.Shapes(piC.name).TopLeftCell.Row + 25
    wsOut.PageSetup.PrintArea = wsOut.Range("A1:N" & lastRow).Address

    names.Add wsOut.name
    Debug.Print "AddCountryMonthlyChartPage [" & country & "]: page added as " & wsOut.name

    ' Cleanup scratch chart + temp sheet
    On Error Resume Next
    co.Delete
    Application.DisplayAlerts = False
    wsScratch.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Exit Sub

EH:
    Debug.Print "AddCountryMonthlyChartPage [" & country & "] FAILED at: " & gStep & " - " & Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Application.DisplayAlerts = False
    If Not wsScratch Is Nothing Then wsScratch.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub SortMonthAggregateKeys(ByRef keys() As String, ByVal monthMap As Object)
    On Error GoTo SafeExit
    Dim i As Long, j As Long
    Dim tmp As String
    Dim ri As Variant, rj As Variant

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            ri = monthMap(keys(i))
            rj = monthMap(keys(j))
            If CStr(rj(3)) < CStr(ri(3)) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
SafeExit:
End Sub
