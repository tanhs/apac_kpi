Attribute VB_Name = "modSecurity"
'modSecurity module
Option Explicit

Public Const DEV_MODE As Boolean = False  '<<< set TRUE only during dev
Public Const SYS_STRUCT_PWD As String = "sheetlock"
Public Const SYS_SHEET_PWD  As String = "sheetlock"
Public Const SYS_IT_ADMIN_PWD  As String = "Itlock"

'========================================================
' Workbook structure lock (prevents rename/delete/move sheets)
'========================================================
Public Sub ProtectWorkbookStructure()
    If DEV_MODE Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Protect Password:=SYS_STRUCT_PWD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Public Sub UnprotectWorkbookStructure()
    If DEV_MODE Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=SYS_STRUCT_PWD
    On Error GoTo 0
End Sub

'========================================================
' Protect a sheet for end-users but still allow VBA edits
'========================================================
Public Sub ProtectSheetUIOnly(ByVal ws As Worksheet, Optional ByVal allowFilter As Boolean = True)
    If DEV_MODE Then Exit Sub
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=SYS_SHEET_PWD
    On Error GoTo 0

    On Error Resume Next
    ws.Protect Password:=SYS_SHEET_PWD, _
        DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFiltering:=allowFilter, AllowSorting:=allowFilter, _
        UserInterfaceOnly:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Public Sub UnprotectSheet(ByVal ws As Worksheet)
    If DEV_MODE Then Exit Sub
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=SYS_SHEET_PWD
    On Error GoTo 0
End Sub

'========================================================
' Apply policy: lock structure + protect key sheets
'========================================================
Public Sub ApplyHardening()
    If modSecurity.DEV_MODE Then Exit Sub

    Dim wb As Workbook: Set wb = ThisWorkbook

    ' --- ALWAYS ensure structure is UNPROTECTED before any Visible changes ---
    On Error Resume Next
    wb.Unprotect Password:=modSecurity.SYS_STRUCT_PWD
    On Error GoTo 0

    ' --- Ensure helper sheets are forced to very hidden while structure is unlocked ---
    Dim wsHost As Worksheet
    Set wsHost = GetWs("zz_CacheHost")
    If Not wsHost Is Nothing Then
        UnprotectSheet wsHost
        wsHost.Visible = xlSheetVeryHidden
    End If

    ' --- zz_CH_<country> cache hosts
    Dim wsZZ As Worksheet
    For Each wsZZ In wb.Worksheets
        If LCase$(left$(wsZZ.name, 6)) = "zz_ch_" Then
            UnprotectSheet wsZZ
            wsZZ.Visible = xlSheetVeryHidden
        End If
    Next wsZZ
    
    Dim wsCalc As Worksheet
    Set wsCalc = GetWs("Calc")
    If Not wsCalc Is Nothing Then
        UnprotectSheet wsCalc
        wsCalc.Visible = xlSheetVeryHidden
    End If

    ' --- ADD: delete scratch sheets (Pivots_Scratch, ChScratch_Tmp, Chart1) ---
    Dim scratchNames(2) As String
    scratchNames(0) = "Pivots_Scratch"
    scratchNames(1) = "ChScratch_Tmp"
    scratchNames(2) = "Chart1"
    Application.DisplayAlerts = False
    Dim i As Integer
    For i = 0 To UBound(scratchNames)
        Dim wsDel As Worksheet
        Set wsDel = Nothing
        On Error Resume Next
        Set wsDel = wb.Worksheets(scratchNames(i))
        If Not wsDel Is Nothing Then wsDel.Delete
        Dim chDel As Chart
        Set chDel = Nothing
        Set chDel = wb.Charts(scratchNames(i))
        If Not chDel Is Nothing Then chDel.Delete
        On Error GoTo 0
    Next i
    
    ' --- DELETE CS__ temp chart sheets ---
    Dim ii As Integer
    For ii = wb.Worksheets.Count To 1 Step -1
        If LCase$(left$(wb.Worksheets(ii).name, Len("CS__"))) = "cs__" Then
            wb.Worksheets(ii).Delete
        End If
    Next ii

    Application.DisplayAlerts = True
    
    ' --- Now do sheet protections ---
    ProtectDataSheet_Strict GetWs("Data")
    ProtectSheetForUsers_PivotFriendly GetWs("Dashboard")
    ProtectSheetForUsers_PivotFriendly GetWs("Pivots")

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase$(left$(ws.name, 10)) = "dashboard_" Then ProtectSheetForUsers_PivotFriendly ws
        If LCase$(left$(ws.name, 7)) = "pivots_" Then ProtectSheetForUsers_PivotFriendly ws
    Next ws

    ' --- LAST STEP: lock workbook structure (rename/delete/move sheets) ---
    On Error Resume Next
    wb.Protect Password:=modSecurity.SYS_STRUCT_PWD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Private Function GetWs(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetWs = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function
'1) Protect Dashboard and Pivot sheets “pivot-friendly”
'Use AllowUsingPivotTables:=True and don’t lock drawing objects if you want chart type changes.
'========================================================
' Protect a sheet but allow Pivot filtering + chart edits
'========================================================
Public Sub ProtectSheetForUsers_PivotFriendly(ByVal ws As Worksheet)
    If modSecurity.DEV_MODE Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=modSecurity.SYS_SHEET_PWD
    On Error GoTo 0

    ' IMPORTANT:
    ' - DrawingObjects:=False  => user can edit charts (change chart type)
    ' - AllowUsingPivotTables:=True => user can filter/sort pivots
    ' - UserInterfaceOnly:=True => VBA can still modify protected sheet later
    ws.Protect Password:=modSecurity.SYS_SHEET_PWD, _
               Contents:=True, _
               DrawingObjects:=False, _
               Scenarios:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               AllowUsingPivotTables:=True, _
               UserInterfaceOnly:=True

    ' Let them click anywhere (optional)
    ws.EnableSelection = xlNoRestrictions
End Sub
'2) Keep Data sheet protected (no editing)
'Data should stay locked (and usually you do want DrawingObjects:=True there).
Public Sub ProtectDataSheet_Strict(ByVal ws As Worksheet)
    If modSecurity.DEV_MODE Then Exit Sub

    On Error Resume Next
    ws.Unprotect Password:=modSecurity.SYS_SHEET_PWD
    On Error GoTo 0

    ws.Protect Password:=modSecurity.SYS_SHEET_PWD, _
               Contents:=True, _
               DrawingObjects:=True, _
               Scenarios:=True, _
               AllowFiltering:=False, _
               AllowSorting:=False, _
               AllowUsingPivotTables:=False, _
               UserInterfaceOnly:=True

    ws.EnableSelection = xlNoSelection
End Sub

