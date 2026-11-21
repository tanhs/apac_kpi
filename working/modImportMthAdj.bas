Attribute VB_Name = "modImportMthAdj"
'modImportMthAdj module
Option Explicit

Private gStep As String
Private Const ADJ_VERSION As String = "v1"


'========================================================
' PUBLIC ENTRY: called by Dashboard button "Import Prev Month Adj"
'========================================================
Public Sub ImportPrevMonthAdj()
    gStep = "ImportPrevMonthAdj: start"
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    modSecurity.ApplyHardening

    ProgressStart "Import Prev Month Adj", 100
    ProgressSet 2, "Selecting adjustment file..."

    Dim filePath As String
    filePath = PickOneExcelFile("Select Prev Month Adjustment Output File")
    If Len(filePath) = 0 Then
        ProgressFinish "Cancelled"
        GoTo CleanExit
    End If

    ProgressSet 10, "Opening file (read-only)..."

    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)

    Dim wsSrc As Worksheet
    gStep = "ImportPrevMonthAdj: ValidatePrevMonthAdjWorkbook"
    If Not ValidatePrevMonthAdjWorkbook(wbSrc, wsSrc) Then
        Err.Raise vbObjectError + 7800, , "Validation failed."
    End If

    ' read header vars
    Dim countryFull As String, monthName As String, adjDate As Date, adjYear As Long
    countryFull = Trim$(CStr(wsSrc.Range("B1").Value))
    monthName = Trim$(CStr(wsSrc.Range("B2").Value))
    adjYear = CLng(Val(wsSrc.Range("B3").Value))
    adjDate = DateSerial(adjYear, MonthNameToNumber(monthName), 1)
    
    Dim countryCode As String
    countryCode = CountryFullNameToCode(countryFull)

    If Not ConfirmAction("Validation done! Are you sure to proceed?", "Confirm") Then
        GoTo CleanFail
    End If

    gStep = "ImportPrevMonthAdj: RunUnlockAll"
    RunUnlockAll

    ProgressSet 20, "Backup BEFORE..."
    'Dim backupToken As String
    'backupToken = modBackup.backupToken()
    'modBackup.BackupWorkbookCopy "Before", backupToken

    ' Apply adjustments into tblData
    ProgressSet 35, "Applying adjustments..."
    gUploadedAt = Now
    gUploadId = BuildAdjUploadId(gUploadedAt)
    
    ApplyMonthAdjustments wsSrc, countryCode, adjDate, gUploadId

    ProgressSet 85, "Refreshing pivots/charts..."
    ThisWorkbook.RefreshAll
    DoEvents

    ProgressSet 90, "Saving workbook..."
    'ThisWorkbook.Save
    'If Not WaitForWorkbookSaved(ThisWorkbook, 30) Then
    '    Err.Raise vbObjectError + 7810, , "Save did not complete within timeout."
    'End If
    'If Not WaitForFileStable(ThisWorkbook.FullName, 5, 60) Then
    '    Err.Raise vbObjectError + 7811, , "Workbook file is still changing on disk (OneDrive may be busy)."
    'End If

    ProgressSet 94, "Backup AFTER..."
    'modBackup.BackupWorkbookCopy "After", backupToken

    ProgressSet 96, "Generating new PDF..."
    ExportAllCharts_Fast_ChartSheets_Direct

    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
    
    ProgressFinish "Completed"

CleanExit:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    ProgressForceClose

    If gRunUnlocked Then RunReharden
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Dim failStep As String
    Dim failNo As Long
    Dim failDesc As String
    failStep = gStep
    failNo = Err.Number
    failDesc = Err.Description

    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    ProgressForceClose
    If gRunUnlocked Then RunReharden

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Import Prev Month Adj failed at step:" & vbCrLf & _
           failStep & vbCrLf & vbCrLf & _
           failNo & " - " & failDesc & vbCrLf & vbCrLf & _
           "All adjustment are not updated. Fix the source file to upload again!", vbCritical
End Sub

'========================================================
' VALIDATION: match the screenshot format (ToUpload style)
'========================================================
Private Function ValidatePrevMonthAdjWorkbook(ByVal wbSrc As Workbook, ByRef outWs As Worksheet) As Boolean
    gStep = "ValidatePrevMonthAdjWorkbook: start"
    On Error GoTo EH

    ' 1) find a sheet that matches the expected layout
    gStep = "ValidatePrevMonthAdjWorkbook: locate ToUpload-like sheet"
    Set outWs = FindPrevMonthAdjSheet(wbSrc)
    If outWs Is Nothing Then
        Err.Raise vbObjectError + 7802, , _
            "Cannot find a worksheet matching the Prev Month Adj template." & vbCrLf & _
            "Expected: A1=Country, A2=Month, A3=Year, A4=Version, A6=APACCentreName, B6=Actual." & vbCrLf & _
            "Sheets found: " & ListSheetNames(wbSrc)
    End If

    ' 2) validate headers
    gStep = "ValidatePrevMonthAdjWorkbook: validate header cells"
    gStep = "ValidatePrevMonthAdjWorkbook: validate A1 header"
    If LCase$(Trim$(CStr(outWs.Range("A1").Value))) <> "country" Then
        Err.Raise vbObjectError + 7803, , "Invalid template: A1 must be 'Country'."
    End If
    gStep = "ValidatePrevMonthAdjWorkbook: validate A2 header"
    If LCase$(Trim$(CStr(outWs.Range("A2").Value))) <> "month" Then
        Err.Raise vbObjectError + 7804, , "Invalid template: A2 must be 'Month'."
    End If

    gStep = "ValidatePrevMonthAdjWorkbook: validate B1 country value"
    Dim countryFull As String
    countryFull = Trim$(CStr(outWs.Range("B1").Value))
    If Len(countryFull) = 0 Then
        Err.Raise vbObjectError + 7805, , "Invalid template: B1 (Country value) is blank."
    End If

    gStep = "ValidatePrevMonthAdjWorkbook: validate B1 country mapping"
    ' Validate country text against the canonical mapper (raises if unknown).
    Dim countryCode As String
    countryCode = CountryFullNameToCode(countryFull)

    gStep = "ValidatePrevMonthAdjWorkbook: validate B2 month name"
    Dim monthName As String
    monthName = Trim$(CStr(outWs.Range("B2").Value))
    If Not IsValidMonthName(monthName) Then
        Err.Raise vbObjectError + 7806, , "Invalid template: B2 must be a full month name (January..December)."
    End If

    gStep = "ValidatePrevMonthAdjWorkbook: validate B3 year"
    Dim adjYear As Long
    If Not IsNumeric(outWs.Range("B3").Value) Then
        Err.Raise vbObjectError + 7815, , "Invalid template: B3 must be numeric year."
    End If
    adjYear = CLng(Val(outWs.Range("B3").Value))

    gStep = "ValidatePrevMonthAdjWorkbook: validate month is previous month"
    Dim adjMonthStart As Date
    Dim currMonthStart  As Date
    adjMonthStart = DateSerial(adjYear, MonthNameToNumber(monthName), 1)
    currMonthStart = DateSerial(Year(Date), Month(Date), 1)
    
    If adjMonthStart >= currMonthStart Then
        Err.Raise vbObjectError + 7816, , _
            "Invalid template: month/year cannot be current or future month."
    End If

    gStep = "ValidatePrevMonthAdjWorkbook: validate B4 version"
    Dim adjVersionControl As String
    adjVersionControl = LCase$(Trim$(CStr(outWs.Range("B4").Value)))
    If adjVersionControl <> ADJ_VERSION Then
        Err.Raise vbObjectError + 7806, , "Invalid template: B4 wrong version of Adjustment template file"
    End If


    gStep = "ValidatePrevMonthAdjWorkbook: validate table headers row 6"
    If LCase$(Trim$(CStr(outWs.Range("A6").Value))) <> LCase$("APACCentreName") Then
        Err.Raise vbObjectError + 7807, , "Invalid template: A6 must be 'APACCentreName'."
    End If
    If LCase$(Trim$(CStr(outWs.Range("B6").Value))) <> "actual" Then
        Err.Raise vbObjectError + 7808, , "Invalid template: B6 must be 'Actual'."
    End If

    ' 3) validate data rows (row 7 onward)
    gStep = "ValidatePrevMonthAdjWorkbook: validate rows 7+ (non-fatal)"
    Dim lastRow As Long
    lastRow = outWs.Cells(outWs.Rows.Count, "A").End(xlUp).Row
    If lastRow < 7 Then
        Err.Raise vbObjectError + 7809, , "Invalid template: no data rows found (expected row 7 onward)."
    End If
    
    Dim r As Long
    Dim errMsg As String
    errMsg = ""
    
    For r = 7 To lastRow
        Dim centerName As String
        centerName = Trim$(CStr(outWs.Cells(r, "A").Value))
    
        Dim vActual As Variant
        vActual = outWs.Cells(r, "B").Value
    
        ' count only usable rows
        If Len(Trim$(centerName)) = 0 Then
            errMsg = errMsg & "Blank center name at " & r & vbCrLf
        End If
        
        If Not IsNumeric(vActual) Or vActual = 0 Then
            errMsg = errMsg & "Actual must be number and not zero at " & r & vbCrLf
        End If
        
    Next r
    
    If Len(errMsg) > 0 Then
        Err.Raise vbObjectError + 7814, , "Invalid row/s: no usable rows found (need Centre + numeric Actual)." & vbCrLf & errMsg
    End If

    gStep = "ValidatePrevMonthAdjWorkbook: validate Country+Center exists in Data for month"
    ValidateAdjCoverageInData outWs, countryCode, adjMonthStart

    ValidatePrevMonthAdjWorkbook = True
    Exit Function

EH:
    Err.Raise vbObjectError + 7899, , gStep & ": " & Err.Description
End Function

Private Function BuildAdjUploadId(ByVal uploadedAt As Date) As String
    Dim baseId As String
    baseId = BuildUploadId(uploadedAt) ' e.g. UPL_20260311_214637_875_6462

    If UCase$(left$(baseId, 4)) = "UPL_" Then
        BuildAdjUploadId = "ADJ_" & Mid$(baseId, 5)
    Else
        BuildAdjUploadId = "ADJ_" & baseId
    End If
End Function

Private Sub ValidateAdjCoverageInData(ByVal wsSrc As Worksheet, ByVal countryCode As String, ByVal adjMonthStart As Date)
    gStep = "ValidateAdjCoverageInData: start"
    On Error GoTo EH

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")

    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 7817, , "Data!tblData has no rows. Cannot validate adjustment coverage."
    End If

    Dim iCountry As Long, iCenter As Long, iDate As Long
    iCountry = lo.ListColumns("Country").Index
    iCenter = lo.ListColumns("Center").Index
    ' Use Date as source-of-truth and derive month from it; Month column may be display/derived.
    iDate = lo.ListColumns("Date").Index

    Dim targetYear As Long, targetMonth As Long
    targetYear = Year(adjMonthStart)
    targetMonth = Month(adjMonthStart)

    ' Build lookup of centers that exist for this country + month in tblData.
    Dim centerMap As Object
    Set centerMap = CreateObject("Scripting.Dictionary")
    centerMap.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim rw As Range
        Set rw = lo.ListRows(i).Range

        If StrComp(Trim$(CStr(rw.Cells(1, iCountry).Value)), countryCode, vbTextCompare) = 0 Then
            Dim vD As Variant
            vD = rw.Cells(1, iDate).Value
            If IsDate(vD) Then
                Dim dRec As Date
                dRec = CDate(vD)
                If Year(dRec) = targetYear And Month(dRec) = targetMonth Then
                    Dim cKey As String
                    cKey = LCase$(Trim$(CStr(rw.Cells(1, iCenter).Value)))
                    If Len(cKey) > 0 Then centerMap(cKey) = True
                End If
            End If
        End If
    Next i

    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    Dim missingCount As Long
    Dim missingMsg As String
    Dim srcRow As Long
    For srcRow = 7 To lastRow
        Dim srcCenter As String
        srcCenter = Trim$(CStr(wsSrc.Cells(srcRow, "A").Value))
        If Len(srcCenter) = 0 Then GoTo NextSrcRow

        If Not centerMap.Exists(LCase$(srcCenter)) Then
            missingCount = missingCount + 1
            If Len(missingMsg) < 2000 Then
                missingMsg = missingMsg & "Row " & srcRow & " [" & srcCenter & "]" & vbCrLf
            End If
        End If
NextSrcRow:
    Next srcRow

    If missingCount > 0 Then
        Err.Raise vbObjectError + 7818, , _
            "Adjustment centers not found in Data!tblData for country [" & countryCode & _
            "] month [" & Format$(adjMonthStart, "mmm-yyyy") & "]. Missing rows: " & missingCount & vbCrLf & missingMsg
    End If
    Exit Sub

EH:
    Err.Raise vbObjectError + 7819, , "ValidateAdjCoverageInData: " & Err.Description
End Sub

'========================================================
' Find the ToUpload-like sheet by structure, not by name
'========================================================
Private Function FindPrevMonthAdjSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If IsPrevMonthAdjSheet(ws) Then
            Set FindPrevMonthAdjSheet = ws
            Exit Function
        End If
    Next ws
    Set FindPrevMonthAdjSheet = Nothing
End Function

Private Function IsPrevMonthAdjSheet(ByVal ws As Worksheet) As Boolean
    On Error GoTo SafeExit

    Dim a1 As String, a2 As String, a6 As String, b6 As String
    a1 = LCase$(Trim$(CStr(ws.Range("A1").Value)))
    a2 = LCase$(Trim$(CStr(ws.Range("A2").Value)))
    a6 = LCase$(Trim$(CStr(ws.Range("A6").Value)))
    b6 = LCase$(Trim$(CStr(ws.Range("B6").Value)))

    If a1 = "country" And a2 = "month" And a6 = LCase$("APACCentreName") And b6 = "actual" Then
        IsPrevMonthAdjSheet = True
        Exit Function
    End If

SafeExit:
    IsPrevMonthAdjSheet = False
End Function

'========================================================
' Apply adjustments into tblData (uses your existing helpers)
'========================================================
Private Sub ApplyMonthAdjustments(ByVal wsSrc As Worksheet, ByVal countryCode As String, ByVal adjDate As Date, gUploadId As String)
    gStep = "ApplyMonthAdjustments: start"
    On Error GoTo EH

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Data")

    Dim lo As ListObject
    Set lo = wsData.ListObjects("tblData")

    EnsureTblDataSchema wsData, lo

    ' This routine assumes run-level unlock has been applied by caller.
    If wsData.ProtectContents Or wsData.ProtectDrawingObjects Or wsData.ProtectScenarios Then
        Err.Raise vbObjectError + 7853, , "Data sheet is protected. RunUnlockAll must be called before ApplyMonthAdjustments."
    End If

    Dim keyIndex As Object
    Set keyIndex = BuildExistingKeyIndex(lo)

    Dim wsAdjLog As Worksheet
    Dim loAdjLog As ListObject
    EnsureAdjLogTable wsData, lo, wsAdjLog, loAdjLog


    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    ' --- skip logging ---
    Dim skipLog As String
    Dim skipCount As Long
    skipLog = ""
    skipCount = 0

    Dim r As Long
    For r = 7 To lastRow
        ' Update adjustment start here
        ApplyOneAdjustmentRow wsSrc, r, lo, loAdjLog, keyIndex, countryCode, adjDate, skipLog, skipCount
NextR:
    Next r

    ' --- REFRESH PATH (subsequent imports) ---
    ' Pivots + charts already exist. Just re-stage tblData,
    ' refresh pivot caches, update static ranges, update charts.
    ' No delete, no recreate - combo layout is permanent.
    gStep = "ImportAndRefresh: refresh pivot caches"
    'ProgressSet 70, "Refreshing pivots..."
    RefreshAllPivotCaches

    If Not MainPivotSourceIsCurrent() Then
        gStep = "ImportAndRefresh: self-heal main pivots"
        'ProgressSet 78, "Rebuilding main pivots..."
        EnsurePivotsSlicersChartsAndTables
    End If

    gStep = "ImportAndRefresh: clear main slicers"
    ClearMainDashboardSlicerFilters

    gStep = "ImportAndRefresh: refresh static ranges + charts"
    'ProgressSet 85, "Refreshing charts..."
    RefreshDashboardCharts True
    RepositionDashboardCharts
    EnsureCountryDashboardCharts

    gStep = "ImportAndRefresh: BuildUploadSummary"
    BuildUploadSummary



CleanExit:
'Debug.Print "Clean Exit "
    ' show skip summary (single msgbox)
    If skipCount > 0 Then
        MsgBox "Prev Month Adj import completed with skipped rows: " & skipCount & vbCrLf & vbCrLf & _
               skipLog, vbExclamation, "Skipped Rows"
    End If

    Exit Sub

EH:
'Debug.Print "EH "
    Err.Raise vbObjectError + 7850, , gStep & ": " & Err.Description
End Sub

Private Sub ApplyOneAdjustmentRow(ByVal wsSrc As Worksheet, ByVal srcRow As Long, _
                                  ByVal loData As ListObject, ByVal loAdjLog As ListObject, _
                                  ByRef keyIndex As Object, _
                                  ByVal countryCode As String, ByVal adjDate As Date, _
                                  ByRef skipLog As String, ByRef skipCount As Long)
    gStep = "ApplyOneAdjustmentRow: start row " & srcRow

    gStep = "ApplyOneAdjustmentRow: read center row " & srcRow
    Dim centerName As String
    centerName = Trim$(CStr(wsSrc.Cells(srcRow, "A").Value))
    If Len(centerName) = 0 Then
        AppendSkip skipLog, skipCount, "Row " & srcRow & ": center is blank."
        Exit Sub
    End If

    gStep = "ApplyOneAdjustmentRow: read adjustment actual row " & srcRow & " center [" & centerName & "]"
    Dim vAdjActual As Variant
    vAdjActual = wsSrc.Cells(srcRow, "B").Value
    If Not IsNumeric(vAdjActual) Then
        AppendSkip skipLog, skipCount, "Row " & srcRow & " [" & centerName & "]: Actual is not numeric."
        Exit Sub
    End If

    Dim adjActual As Double
    adjActual = CDbl(vAdjActual)

    Dim iCountry As Long, iRegion As Long, iCenter As Long, iDate As Long, iBudget As Long
    iCountry = loData.ListColumns("Country").Index
    iRegion = loData.ListColumns("Region").Index
    iCenter = loData.ListColumns("Center").Index
    ' Use Date as source-of-truth and derive month from it; Month column may be display/derived.
    iDate = loData.ListColumns("Date").Index
    iBudget = loData.ListColumns("Budget").Index

    Dim toDelete As Collection
    Set toDelete = New Collection

    Dim totBudget As Double
    Dim adjRegion As String

    Dim targetYear As Long
    Dim targetMonth As Long
    targetYear = Year(adjDate)
    targetMonth = Month(adjDate)

    Dim i As Long
    gStep = "ApplyOneAdjustmentRow: scan tblData matches row " & srcRow & " center [" & centerName & "]"
    If Not loData.DataBodyRange Is Nothing Then
        For i = 1 To loData.ListRows.Count
            Dim rw As Range
            Set rw = loData.ListRows(i).Range

            ' Match candidate rows by business key for adjustment: Country + Center.
            ' Region is intentionally excluded from this match.
            If StrComp(Trim$(CStr(rw.Cells(1, iCountry).Value)), countryCode, vbTextCompare) = 0 And _
               StrComp(Trim$(CStr(rw.Cells(1, iCenter).Value)), centerName, vbTextCompare) = 0 Then

                Dim vD As Variant
                vD = rw.Cells(1, iDate).Value
                If IsDate(vD) Then
                    ' Keep rows in the same adjustment month (not exact day).
                    Dim dRec As Date
                    dRec = CDate(vD)
                    If Year(dRec) = targetYear And Month(dRec) = targetMonth Then
                        toDelete.Add i

                        If IsNumeric(rw.Cells(1, iBudget).Value) Then
                            totBudget = totBudget + CDbl(rw.Cells(1, iBudget).Value)
                        End If

                        ' Capture region from the first month-matched row so upsert uses original region.
                        If Len(adjRegion) = 0 Then
                            adjRegion = Trim$(CStr(rw.Cells(1, iRegion).Value))
                        End If
                    End If
                End If
            End If
        Next i
    End If

    gStep = "ApplyOneAdjustmentRow: resolve region row " & srcRow & " center [" & centerName & "]"
    If Len(adjRegion) = 0 Then
        adjRegion = LookupRegionByCenter(loData, centerName)
    End If

    If toDelete.Count = 0 Then
        AppendSkip skipLog, skipCount, "Row " & srcRow & " [" & centerName & "]: no matching data found to adjust."
        Exit Sub
    End If

    If Len(adjRegion) = 0 Then
        AppendSkip skipLog, skipCount, "Row " & srcRow & " [" & centerName & "]: region cannot be determined."
        Exit Sub
    End If

    ' Move old rows to AdjLog then delete from Data (delete from bottom to top).
    gStep = "ApplyOneAdjustmentRow: move old rows to AdjLog row " & srcRow & " center [" & centerName & "]"
    For i = toDelete.Count To 1 Step -1
        Dim rowIdx As Long
        rowIdx = CLng(toDelete(i))

        Dim rowVals As Variant
        rowVals = loData.ListRows(rowIdx).Range.Value

        Dim lrOld As ListRow
        Set lrOld = loAdjLog.ListRows.Add
        lrOld.Range.Cells(1, 1).Resize(1, UBound(rowVals, 2)).Value = rowVals

        loData.ListRows(rowIdx).Delete
    Next i

    ' Key index row numbers are stale after deletes -> rebuild before upsert.
    gStep = "ApplyOneAdjustmentRow: rebuild key index row " & srcRow & " center [" & centerName & "]"
    Set keyIndex = BuildExistingKeyIndex(loData)

    Dim mStart As Date
    mStart = DateSerial(Year(adjDate), Month(adjDate), 1)

    Dim variance As Double
    variance = adjActual - totBudget

    Dim ach As Variant
    If totBudget = 0 Then
        ach = Empty
    Else
        ach = adjActual / totBudget
    End If

    gStep = "ApplyOneAdjustmentRow: upsert adjusted row " & srcRow & " center [" & centerName & "]"
    UpsertRow loData, keyIndex, countryCode, adjRegion, centerName, adjDate, mStart, _
              Format$(mStart, "mmm-yyyy"), totBudget, adjActual, variance, ach
End Sub

Private Sub EnsureAdjLogTable(ByVal wsData As Worksheet, ByVal loData As ListObject, _
                              ByRef wsAdjLog As Worksheet, ByRef loAdjLog As ListObject)
    On Error GoTo EH

    On Error Resume Next
    Set wsAdjLog = ThisWorkbook.Worksheets("AdjLog")
    On Error GoTo 0

    If wsAdjLog Is Nothing Then
        Set wsAdjLog = ThisWorkbook.Worksheets.Add(After:=wsData)
        wsAdjLog.name = "AdjLog"
    End If

    On Error Resume Next
    Set loAdjLog = wsAdjLog.ListObjects("tblDataOld")
    On Error GoTo 0

    If loAdjLog Is Nothing Then
        wsAdjLog.Cells.Clear
        wsAdjLog.Range("A1").Resize(1, loData.ListColumns.Count).Value = loData.HeaderRowRange.Value
        wsAdjLog.Range("A2").Resize(1, loData.ListColumns.Count).ClearContents
        Set loAdjLog = wsAdjLog.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=wsAdjLog.Range("A1").Resize(2, loData.ListColumns.Count), _
            XlListObjectHasHeaders:=xlYes)
        loAdjLog.name = "tblDataOld"
        If Not loAdjLog.DataBodyRange Is Nothing Then
            If Application.WorksheetFunction.CountA(loAdjLog.DataBodyRange.Rows(1)) = 0 Then
                loAdjLog.ListRows(1).Delete
            End If
        End If
    Else
        If loAdjLog.ListColumns.Count <> loData.ListColumns.Count Then
            loAdjLog.Resize loAdjLog.Range.Resize(loAdjLog.Range.Rows.Count, loData.ListColumns.Count)
        End If
        loAdjLog.HeaderRowRange.Resize(1, loData.ListColumns.Count).Value = loData.HeaderRowRange.Value
    End If

    ' Keep Month/MonthName formatting aligned with Data!tblData.
    On Error Resume Next
    loAdjLog.ListColumns("Month").Range.NumberFormat = loData.ListColumns("Month").Range.NumberFormat
    loAdjLog.ListColumns("MonthName").Range.NumberFormat = loData.ListColumns("MonthName").Range.NumberFormat
    On Error GoTo EH

    Exit Sub

EH:
    Err.Raise vbObjectError + 7860, , "EnsureAdjLogTable: " & Err.Description
End Sub
'========================================================
' tblData helpers (filter-less loops to avoid pivot/table filter side-effects)
'========================================================
Private Function LookupRegionByCenter(ByVal lo As ListObject, ByVal centerName As String) As String
    Dim r As Long
    If lo.DataBodyRange Is Nothing Then Exit Function

    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("Center").Index).Value)), centerName, vbTextCompare) = 0 Then
            LookupRegionByCenter = Trim$(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns("Region").Index).Value))
            Exit Function
        End If
    Next r
    LookupRegionByCenter = ""
End Function

Private Function SumActualForMonth(ByVal lo As ListObject, ByVal countryCode As String, ByVal centerName As String, ByVal dStart As Date, ByVal dEnd As Date) As Double
    Dim r As Long
    Dim sum As Double
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim iCountry As Long, iCenter As Long, iDate As Long, iActual As Long
    iCountry = lo.ListColumns("Country").Index
    iCenter = lo.ListColumns("Center").Index
    iDate = lo.ListColumns("Date").Index
    iActual = lo.ListColumns("Actual").Index

    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, iCountry).Value)), countryCode, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, iCenter).Value)), centerName, vbTextCompare) = 0 Then
                Dim vD As Variant
                vD = lo.DataBodyRange.Cells(r, iDate).Value
                If IsDate(vD) Then
                    Dim d As Date
                    d = CDate(vD)
                    If d >= dStart And d <= dEnd Then
                        If IsNumeric(lo.DataBodyRange.Cells(r, iActual).Value) Then
                            sum = sum + CDbl(lo.DataBodyRange.Cells(r, iActual).Value)
                        End If
                    End If
                End If
            End If
        End If
    Next r

    SumActualForMonth = sum
End Function

'========================================================
' Country + month parsing helpers
'========================================================
Private Function CountryFullNameToCode(ByVal countryFull As String) As String
    Dim s As String
    s = LCase$(Trim$(countryFull))

    Select Case s
        Case "malaysia":   CountryFullNameToCode = "MY"
        Case "japan":      CountryFullNameToCode = "JP"
        Case "china":      CountryFullNameToCode = "CN"
        Case "singapore":  CountryFullNameToCode = "SG"
        Case Else
            Err.Raise vbObjectError + 7820, , "Unknown country name: [" & countryFull & "]"
    End Select
End Function

Private Function IsValidMonthName(ByVal monthName As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(monthName))
    Select Case s
        Case "january", "february", "march", "april", "may", "june", _
             "july", "august", "september", "october", "november", "december"
            IsValidMonthName = True
        Case Else
            IsValidMonthName = False
    End Select
End Function


Private Function MonthNameToNumber(ByVal monthName As String) As Long
    Dim s As String
    s = LCase$(Trim$(monthName))

    Select Case s
        Case "january": MonthNameToNumber = 1
        Case "february": MonthNameToNumber = 2
        Case "march": MonthNameToNumber = 3
        Case "april": MonthNameToNumber = 4
        Case "may": MonthNameToNumber = 5
        Case "june": MonthNameToNumber = 6
        Case "july": MonthNameToNumber = 7
        Case "august": MonthNameToNumber = 8
        Case "september": MonthNameToNumber = 9
        Case "october": MonthNameToNumber = 10
        Case "november": MonthNameToNumber = 11
        Case "december": MonthNameToNumber = 12
        Case Else
            Err.Raise vbObjectError + 7821, , "Invalid month name: [" & monthName & "]"
    End Select
End Function

'========================================================
' File picker
'========================================================
Private Function PickOneExcelFile(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx; *.xlsm; *.xls"
        If .Show <> -1 Then
            PickOneExcelFile = ""
            Exit Function
        End If
        PickOneExcelFile = .selectedItems(1)
    End With
End Function

Private Function ListSheetNames(ByVal wb As Workbook) As String
    Dim ws As Worksheet, s As String
    For Each ws In wb.Worksheets
        s = s & ws.name & ", "
    Next ws
    If Len(s) >= 2 Then s = left$(s, Len(s) - 2)
    ListSheetNames = s
End Function

Private Sub AppendSkip(ByRef logText As String, ByRef cnt As Long, ByVal msg As String)
    Const MAX_CHARS As Long = 2500 ' keep msgbox readable
    cnt = cnt + 1

    If Len(logText) < MAX_CHARS Then
        logText = logText & msg & vbCrLf
        If Len(logText) >= MAX_CHARS Then
            logText = logText & "... (more skipped rows not shown)" & vbCrLf
        End If
    End If
End Sub
