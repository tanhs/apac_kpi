Attribute VB_Name = "modImportValidation"
Option Explicit

'=========================================================
' Import Validation (Pre-Import, Full Scan)
'
' Contract:
'   - ValidateSelectedWorkbooks(...) returns True only if ALL checks pass
'   - If ANY sheet fails -> returns False with consolidated MsgBox
'
' Caller decides CleanFail handling.
'=========================================================

Private Const MAX_MSG_CHARS As Long = 30000 ' MsgBox practical limit-ish


'-------------------------------------------------------------------------------
' Validates all selected workbook file paths (FileDialog.SelectedItems).
'
' @param selectedItems FileDialog.SelectedItems collection
' @param ByRef gStep A debug step string from caller (Main module)
' @return True if ALL sheets in ALL workbooks pass validation; else False
'-------------------------------------------------------------------------------
Public Function ValidateSelectedWorkbooks(ByVal selectedItems As Variant, ByRef gStep As String) As Boolean
    gStep = "modImportValidation.ValidateSelectedWorkbooks: start"

    Dim errors As Collection
    Set errors = New Collection

    Dim i As Long
    For i = 1 To selectedItems.Count
        Dim filePath As String
        filePath = CStr(selectedItems(i))

        gStep = "modImportValidation.ValidateSelectedWorkbooks: ValidateOneWorkbook " & filePath
        ValidateOneWorkbook filePath, errors, gStep
    Next i

    If errors.Count > 0 Then
        gStep = "modImportValidation.ValidateSelectedWorkbooks: show consolidated errors"
        ShowConsolidatedErrors errors
        ValidateSelectedWorkbooks = False
        Exit Function
    End If

    ValidateSelectedWorkbooks = True
End Function


'-------------------------------------------------------------------------------
' Validate one workbook (open read-only) and validate all worksheets.
' Adds all failures into errors collection; does NOT raise.
'-------------------------------------------------------------------------------
Private Sub ValidateOneWorkbook(ByVal filePath As String, ByRef errors As Collection, ByRef gStep As String)
    gStep = "modImportValidation.ValidateOneWorkbook: open " & filePath

    Dim wb As Workbook
    On Error GoTo OPEN_FAIL

    Set wb = OpenWorkbookReadOnly_Safe(filePath)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Type = xlWorksheet Then
            gStep = "modImportValidation.ValidateOneWorkbook: ValidateOneWorksheet [" & wb.name & "!" & ws.name & "]"
            ValidateOneWorksheet wb, ws, errors, gStep
        End If
    Next ws

CLEAN_EXIT:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

OPEN_FAIL:
    errors.Add "File: " & filePath & vbCrLf & _
               "? Cannot open workbook: " & Err.Number & " - " & Err.Description & vbCrLf & _
               String(60, "-")
    Resume CLEAN_EXIT
End Sub

'-------------------------------------------------------------------------------
' Validate one worksheet against all rules.
' Adds failures into errors collection.
'-------------------------------------------------------------------------------
Private Sub ValidateOneWorksheet(ByVal wb As Workbook, ByVal ws As Worksheet, ByRef errors As Collection, ByRef gStep As String)
    Dim hdr As String
    hdr = "Workbook: " & wb.name & vbCrLf & _
          "Worksheet: " & ws.name & vbCrLf

    '===========================
    ' Rule 8: Worksheet naming convention
    ' - exactly 3 parts separated by "-"
    ' - part1: MY/JP/SG/CN only
    ' - part2: alphanumeric
    ' - part3: alphanumeric
    ' Total char should not exceed 30 including dash and space
    '===========================
    gStep = "modImportValidation.ValidateOneWorksheet: Rule8 sheet name [" & wb.name & "!" & ws.name & "]"
    Dim reason As String
    If Not ValidateSheetNameDetailed(ws.name, reason) Then
        AddError errors, hdr, "? Invalid sheet name: " & ws.name & vbCrLf & reason
    End If

    '===========================
    ' Rule 5: Month must be at C3 and must have data
    ' Rule 1: Month cell C3 must be full month name Jan..Dec
    '===========================
    gStep = "modImportValidation.ValidateOneWorksheet: Rule1/5 month C3 [" & wb.name & "!" & ws.name & "]"
    Dim vMonth As String
    vMonth = Trim$(CStr(ws.Range("C3").Value))

    If Len(vMonth) = 0 Then
        AddError errors, hdr, "? Month is missing. Cell C3 must contain a month (January..December)."
    ElseIf Not IsFullMonthName(vMonth) Then
        AddError errors, hdr, "? Invalid Month at C3: [" & vMonth & "]. Must be full month name (January..December)."
    End If

    '===========================
    ' Rule 6: Headers fixed at row 4
    ' Day at C4, Budget at D4, Actual at E4, Variance at F4
    '===========================
    gStep = "modImportValidation.ValidateOneWorksheet: Rule6 headers row4 [" & wb.name & "!" & ws.name & "]"
    If LCase$(Trim$(CStr(ws.Range("C4").Value))) <> "day" Then
        AddError errors, hdr, "? Header mismatch at C4. Expected 'Day'. Found [" & CStr(ws.Range("C4").Value) & "]."
    End If
    If LCase$(Trim$(CStr(ws.Range("D4").Value))) <> "budget" Then
        AddError errors, hdr, "? Header mismatch at D4. Expected 'Budget'. Found [" & CStr(ws.Range("D4").Value) & "]."
    End If
    If LCase$(Trim$(CStr(ws.Range("E4").Value))) <> "actual" Then
        AddError errors, hdr, "? Header mismatch at E4. Expected 'Actual'. Found [" & CStr(ws.Range("E4").Value) & "]."
    End If
    If LCase$(Trim$(CStr(ws.Range("F4").Value))) <> "variance" Then
        AddError errors, hdr, "? Header mismatch at F4. Expected 'Variance'. Found [" & CStr(ws.Range("F4").Value) & "]."
    End If

    ' If headers are wrong, continuing row checks can produce confusing errors.
    ' Still, you asked "complete all validation" – we continue.

    '===========================
    ' Data rows scan
    ' Rule 2: Day column C must be yyyy-mm-dd AND not allow format change
    ' Rule 3: Budget/Actual numeric and 2 decimals format ####.00
    ' Rule 4: Variance must be Budget-Actual (not manually tallied)
    '===========================
    gStep = "modImportValidation.ValidateOneWorksheet: scan rows [" & wb.name & "!" & ws.name & "]"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 5 Then Exit Sub

    Dim r As Long
    For r = 5 To lastRow
        Dim hasAny As Boolean
        hasAny = (Len(Trim$(CStr(ws.Cells(r, "C").Value))) > 0) _
                 Or (Len(Trim$(CStr(ws.Cells(r, "D").Value))) > 0) _
                 Or (Len(Trim$(CStr(ws.Cells(r, "E").Value))) > 0) _
                 Or (Len(Trim$(CStr(ws.Cells(r, "F").Value))) > 0)

        If Not hasAny Then GoTo nextRow

        ' Day check
        Dim dayVal As Variant
        dayVal = ws.Cells(r, "C").Value

        If Not IsDate(dayVal) Then
            AddError errors, hdr, "? Row " & r & " Day (C" & r & ") is not a valid date."
        Else
            ' enforce number format EXACT "yyyy-mm-dd"
            Dim fmt As String
            fmt = LCase$(Trim$(CStr(ws.Cells(r, "C").NumberFormat)))
            If fmt <> "yyyy-mm-dd" Then
                AddError errors, hdr, "? Row " & r & " Day format invalid at C" & r & ". Expected NumberFormat 'yyyy-mm-dd'. Found [" & ws.Cells(r, "C").NumberFormat & "]."
            End If
        End If

        ' Budget / Actual numeric + 2 decimals format
        Dim bVal As Variant, aVal As Variant
        bVal = ws.Cells(r, "D").Value
        aVal = ws.Cells(r, "E").Value

        If Not IsNumericOrBlank(bVal) Then
            AddError errors, hdr, "? Row " & r & " Budget (D" & r & ") must be numeric."
        ElseIf Not IsBlankValue(bVal) Then
            ' Rule 9: Budget must be > 0 (prevents #DIV/0! in Ach% calculated field)
            If CDbl(bVal) <= 0 Then
                AddError errors, hdr, "? Row " & r & " Budget (D" & r & ") must be greater than 0. Found [" & bVal & "]."
            End If
            If Not HasTwoDecimalFormat(ws.Cells(r, "D").NumberFormat) Then
                AddError errors, hdr, "? Row " & r & " Budget format invalid at D" & r & ". Expected 2 decimals (e.g., ####.00). Found [" & ws.Cells(r, "D").NumberFormat & "]."
            End If
        End If

        If Not IsNumericOrBlank(aVal) Then
            AddError errors, hdr, "? Row " & r & " Actual (E" & r & ") must be numeric."
        ElseIf Not IsBlankValue(aVal) Then
            If Not HasTwoDecimalFormat(ws.Cells(r, "E").NumberFormat) Then
                AddError errors, hdr, "? Row " & r & " Actual format invalid at E" & r & ". Expected 2 decimals (e.g., ####.00). Found [" & ws.Cells(r, "E").NumberFormat & "]."
            End If
        End If

        ' Variance check: numeric and equals Budget-Actual (2dp)
        Dim vVal As Variant
        vVal = ws.Cells(r, "F").Value

        If Not IsNumericOrBlank(vVal) Then
            AddError errors, hdr, "? Row " & r & " Variance (F" & r & ") must be numeric."
        Else
            If IsNumeric(bVal) And IsNumeric(aVal) And IsNumeric(vVal) Then
                Dim expV As Double
                expV = Round(CDbl(aVal) - CDbl(bVal), 2)

                If Round(CDbl(vVal), 2) <> expV Then
                    AddError errors, hdr, _
                        "? Row " & r & " Variance mismatch at F" & r & "." & vbCrLf & _
                        "Expected: Actual - Budget = " & Format$(expV, "0.00") & vbCrLf & _
                        "Found: " & Format$(CDbl(vVal), "0.00")
                End If
            End If
        End If

nextRow:
    Next r
End Sub


'=========================================================
' Consolidated error reporting
'=========================================================
Private Sub AddError(ByRef errors As Collection, ByVal header As String, ByVal details As String)
    errors.Add header & details & vbCrLf & String(60, "-")
End Sub

Private Sub ShowConsolidatedErrors(ByVal errors As Collection)
    Dim msg As String
    msg = "Import validation failed." & vbCrLf & _
          "No data was imported." & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To errors.Count
        Dim part As String
        part = CStr(errors(i))

        If Len(msg) + Len(part) + 2 > MAX_MSG_CHARS Then
            msg = msg & vbCrLf & "…(more errors exist; message truncated)..." & vbCrLf
            Exit For
        End If

        msg = msg & part & vbCrLf
    Next i

    MsgBox msg, vbCritical, "Import Validation"
End Sub


'=========================================================
' Helpers
'=========================================================
Private Function IsFullMonthName(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(CStr(s)))
        Case "january", "february", "march", "april", "may", "june", _
             "july", "august", "september", "october", "november", "december"
            IsFullMonthName = True
        Case Else
            IsFullMonthName = False
    End Select
End Function

'========================================================
' Validate worksheet name with explicit reason
'
' Format: Country-Region-Center
' Rules:
'   - Total length <= 30
'   - Country: MY | JP | SG | CN
'   - Region: alphanumeric
'   - Center: alphanumeric
'
' @param sheetName Worksheet name
' @param ByRef reason If invalid, contains human-readable reason
' @return True if valid
'========================================================
Private Function ValidateSheetNameDetailed(ByVal sheetName As String, ByRef reason As String) As Boolean

    reason = ""

    ' Rule: length
    If Len(sheetName) > 30 Then
        reason = "Name exceeds 30 characters (including dashes)."
        Exit Function
    End If

    ' Split parts
    Dim parts() As String
    parts = Split(sheetName, "-")

    If UBound(parts) <> 2 Then
        reason = "Must follow format Country-Region-Center (exactly 2 dashes)."
        Exit Function
    End If

    Dim country As String, region As String, center As String
    country = parts(0)
    region = parts(1)
    center = parts(2)

    ' Country validation
    Select Case country
        Case "MY", "JP", "SG", "CN"
        Case Else
            reason = "Invalid country [" & country & "]. Allowed: MY, JP, SG, CN."
            Exit Function
    End Select

    ' Region
    If Len(region) = 0 Then
        reason = "Region cannot be empty."
        Exit Function
    End If

    If Not IsAlphaNumeric(region) Then
        reason = "Region must be alphanumeric only. Found: [" & region & "]"
        Exit Function
    End If

    ' Center
    If Len(center) = 0 Then
        reason = "Center cannot be empty."
        Exit Function
    End If

    If Not IsAlphaNumeric(center) Then
        reason = "Center must be alphanumeric only. Found: [" & center & "]"
        Exit Function
    End If

    ValidateSheetNameDetailed = True
End Function

Private Function IsBlankValue(ByVal v As Variant) As Boolean
    IsBlankValue = (IsEmpty(v) Or Trim$(CStr(v)) = "" Or v = 0)
End Function

Private Function IsNumericOrBlank(ByVal v As Variant) As Boolean
    If IsBlankValue(v) Then
        IsNumericOrBlank = True
    Else
        IsNumericOrBlank = IsNumeric(v)
    End If
End Function

Private Function HasTwoDecimalFormat(ByVal fmt As String) As Boolean
    ' Accept a few common 2-decimal patterns (Excel may store variations)
    Dim f As String
    f = LCase$(Trim$(fmt))

    ' common:
    ' 0.00
    ' #,##0.00
    ' ####.00 (user asked)
    ' ###0.00 etc
    If InStr(1, f, ".00", vbTextCompare) > 0 Then
        HasTwoDecimalFormat = True
    Else
        HasTwoDecimalFormat = False
    End If
End Function

'--------------------------------------------------
' Opens a workbook read-only with hardened settings and clear error reporting.
'
' @param filePath Full path to workbook
' @return Opened workbook object (never Nothing on success)
'--------------------------------------------------
'--------------------------------------------------
' Opens a workbook read-only with hardened settings and Protected View handling.
'
' @param filePath Full path to workbook
' @return Opened workbook object (never Nothing on success)
'--------------------------------------------------
Private Function OpenWorkbookReadOnly_Safe(ByVal filePath As String) As Workbook

    Dim wb As Workbook
    Dim pvBefore As Long
    Dim pvAfter As Long

    If Len(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 7401, , "Open failed: filePath is empty."
    End If
    If Dir$(filePath) = vbNullString Then
        Err.Raise vbObjectError + 7401, , "Open failed: file not found:" & vbCrLf & filePath
    End If

    ' Snapshot ProtectedViewWindows count BEFORE open
    pvBefore = Application.ProtectedViewWindows.Count

    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevUpdating As Boolean: prevUpdating = Application.ScreenUpdating
    Dim prevSec As MsoAutomationSecurity: prevSec = Application.AutomationSecurity

    On Error GoTo EH

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Err.Clear

    ' Try normal open first (no prompts, no link updates)
    Set wb = Workbooks.Open( _
        fileName:=filePath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False, _
        Notify:=False)

    ' If Workbook is Nothing AND no error -> likely Protected View
    If wb Is Nothing And Err.Number = 0 Then

        pvAfter = Application.ProtectedViewWindows.Count

        If pvAfter > pvBefore Then
            Dim pv As ProtectedViewWindow
            For Each pv In Application.ProtectedViewWindows
                ' Match by full path if possible; fall back to file name
                If LCase$(pv.SourceName) = LCase$(filePath) _
                   Or LCase$(Dir$(pv.SourceName)) = LCase$(Dir$(filePath)) Then

                    ' Attempt to switch to editable workbook (may be blocked by policy)
                    On Error Resume Next
                    pv.Edit
                    On Error GoTo EH

                    ' After Edit, workbook should exist
                    On Error Resume Next
                    Set wb = Workbooks(Dir$(filePath))
                    On Error GoTo EH

                    Exit For
                End If
            Next pv
        End If

        If wb Is Nothing Then
            Err.Raise vbObjectError + 7402, , _
                "Workbook opened in Protected View (no Workbook object created)." & vbCrLf & _
                "File: " & filePath & vbCrLf & vbCrLf & _
                "Fix options:" & vbCrLf & _
                "1) Open the file manually once and click 'Enable Editing', then try again." & vbCrLf & _
                "2) Add the folder as a Trusted Location (Trust Center)." & vbCrLf & _
                "3) Right-click file > Properties > Unblock (if present)."
        End If
    End If

    Set OpenWorkbookReadOnly_Safe = wb

CleanExit:
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    Exit Function

EH:
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description

    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    On Error GoTo 0

    Err.Raise vbObjectError + 7403, , _
        "Workbooks.Open failed." & vbCrLf & _
        "File: " & filePath & vbCrLf & _
        "Err: " & errNo & " - " & errMsg
End Function

Private Function IsAlphaNumeric(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "^[A-Za-z0-9]+$"
    IsAlphaNumeric = re.Test(s)
End Function
