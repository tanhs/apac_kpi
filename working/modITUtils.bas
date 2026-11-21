Attribute VB_Name = "modITUtils"
'modITUtils module

Option Explicit
Private Const FALLBACK_UNKNOWN_USER As String = "UNKNOWN_USER"
'========================================================
' Delete all worksheets except core & country system sheets
'========================================================
Public Sub DeleteAllNonCoreExceptCountrySheets()
    MsgBox "Don't run this", vbExclamation
    Exit Sub

    'If Not PromptPasswordAndValidate(SYS_IT_ADMIN_PWD) Then
    '    MsgBox "Invalid password or cancelled.", vbExclamation
    '    Exit Sub
    'End If
    
    RunUnlockAll

    gStep = "DeleteAllNonCoreExceptCountrySheets: start"
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim ws As Worksheet
    Dim keepSheets As Object

    'Sheets that must remain (Exact matches)
    Set keepSheets = CreateObject("Scripting.Dictionary")
    keepSheets.CompareMode = vbTextCompare

    keepSheets.Add "Data", True
    keepSheets.Add "Pivots", True
    keepSheets.Add "Dashboard", True
    keepSheets.Add "Upload Summary", True
    keepSheets.Add "ITAdmin", True
    keepSheets.Add "zz_CacheHost", True
    

    'Loop backwards to avoid collection issues
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)

        'Check 1: Exact match in Dictionary
        'Check 2: Name starts with "Dashboard_"
        'Check 3: Name starts with "Pivots_"
        If Not keepSheets.Exists(ws.name) And _
           Not ws.name Like "Dashboard_*" And _
           Not ws.name Like "Pivots_*" Or _
            ws.name = "Pivots_Scratch" Then
            
            ' Only delete if it doesn't meet any of the criteria above
            ws.Delete
            
        End If
    Next i

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    RunReharden
    MsgBox "Done", vbExclamation
    Exit Sub

CleanFail:
    MsgBox "Error during sheet cleanup at step: " & gStep & vbCrLf & Err.Description, vbCritical
    Resume CleanExit

End Sub

'========================================================
' Delete all worksheets except core system sheets
'========================================================
Public Sub DeleteAllNonCoreSheets()
    MsgBox "Don't run this", vbExclamation
    Exit Sub
    
    'If Not PromptPasswordAndValidate(SYS_IT_ADMIN_PWD) Then
     '   MsgBox "Invalid password or cancelled.", vbExclamation
      '  Exit Sub
    'End If
    
    RunUnlockAll

    gStep = "DeleteAllNonCoreSheets: start"
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim ws As Worksheet
    Dim keepSheets As Object

    'Sheets that must remain
    Set keepSheets = CreateObject("Scripting.Dictionary")
    keepSheets.CompareMode = vbTextCompare

    keepSheets.Add "Data", True
    keepSheets.Add "Pivots", True
    keepSheets.Add "Dashboard", True
    keepSheets.Add "Upload Summary", True
    keepSheets.Add "ITAdmin", True

    'Loop backwards to avoid collection issues
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1

        Set ws = ThisWorkbook.Worksheets(i)

        If Not keepSheets.Exists(ws.name) Then
            ws.Delete
        End If

    Next i

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    RunReharden
    MsgBox "Done", vbExclamation
    Exit Sub

CleanFail:
    MsgBox "Error during sheet cleanup at step: " & gStep & vbCrLf & Err.Description, vbCritical
    Resume CleanExit

End Sub

'========================================================
' Show masked password dialog and validate
'========================================================
Public Function PromptPasswordAndValidate(ByVal expectedPwd As String) As Boolean

    Dim f As frmPasswordPrompt
    Set f = New frmPasswordPrompt
    
    f.Show vbModal
    
    If f.Cancelled Then
        PromptPasswordAndValidate = False
    Else
        PromptPasswordAndValidate = (StrComp(f.UserPassword, expectedPwd, vbBinaryCompare) = 0)
    End If
    
    Unload f
    Set f = Nothing

End Function
'========================================================
' Ask user confirmation (Yes / No)
'========================================================
Public Function ConfirmAction(ByVal msg As String, Optional ByVal title As String = "Confirm") As Boolean

    Dim result As VbMsgBoxResult
    
    result = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, title)
    
    If result = vbYes Then
        ConfirmAction = True
    Else
        ConfirmAction = False
    End If

End Function


'====================================================
' Returns the best available user name
'====================================================
Public Function GetBestUserName() As String
    
    Dim s As String
    
    '1. Windows login
    s = Trim$(Environ$("USERNAME"))
    If Len(s) > 0 Then
        GetBestUserName = s
        Exit Function
    End If
    
    '2. Excel user
    s = Trim$(Application.UserName)
    If Len(s) > 0 Then
        GetBestUserName = s
        Exit Function
    End If
    
    '3. OneDrive owner hint
    s = Trim$(GetOneDriveOwnerHint())
    If Len(s) > 0 Then
        GetBestUserName = s
        Exit Function
    End If
    
    GetBestUserName = FALLBACK_UNKNOWN_USER
    
End Function

'====================================================
' Returns normalized key for logging
'====================================================
Public Function GetBestUserKey() As String
    
    Dim rawValue As String
    rawValue = GetBestUserName()
    
    GetBestUserKey = NormalizeUserKey(rawValue)

End Function

'====================================================
' Returns OneDrive local sync path
'====================================================
Public Function GetOneDrivePathSafe() As String
    
    Dim s As String
    
    s = Trim$(Environ$("OneDrive"))
    If Len(s) > 0 Then
        GetOneDrivePathSafe = s
        Exit Function
    End If
    
    s = Trim$(Environ$("OneDriveCommercial"))
    If Len(s) > 0 Then
        GetOneDrivePathSafe = s
        Exit Function
    End If
    
    s = Trim$(Environ$("OneDriveConsumer"))
    If Len(s) > 0 Then
        GetOneDrivePathSafe = s
        Exit Function
    End If
    
    GetOneDrivePathSafe = vbNullString

End Function

'====================================================
' Attempt to extract owner name from OneDrive path
'====================================================
Private Function GetOneDriveOwnerHint() As String
    
    Dim path As String
    Dim parts() As String
    Dim i As Long
    
    path = GetOneDrivePathSafe()
    
    If Len(path) = 0 Then Exit Function
    
    parts = Split(path, "\")
    
    For i = LBound(parts) To UBound(parts) - 1
        If LCase$(parts(i)) = "users" Then
            GetOneDriveOwnerHint = parts(i + 1)
            Exit Function
        End If
    Next i

End Function

'====================================================
' Normalize user key for logs
'====================================================
Private Function NormalizeUserKey(ByVal rawValue As String) As String
    
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim resultValue As String
    
    s = Trim$(UCase$(rawValue))
    
    For i = 1 To Len(s)
    
        ch = Mid$(s, i, 1)
        
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            resultValue = resultValue & ch
        Else
            resultValue = resultValue & "_"
        End If
        
    Next i
    
    Do While InStr(resultValue, "__") > 0
        resultValue = Replace(resultValue, "__", "_")
    Loop
    
    NormalizeUserKey = resultValue

End Function

'====================================================
' Debug helper
'====================================================
Public Sub ShowUserIdentitySummary()

    Dim msg As String
    
    msg = msg & "Windows User : " & Environ$("USERNAME") & vbCrLf
    msg = msg & "Excel User   : " & Application.UserName & vbCrLf
    msg = msg & "OneDrive Path: " & GetOneDrivePathSafe() & vbCrLf
    msg = msg & "Best User    : " & GetBestUserName()
    
    MsgBox msg, vbInformation, "User Identity"

End Sub
