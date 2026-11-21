Attribute VB_Name = "modSecurityManager"
'==================================================
' Enables sheet delete from Excel menu by unlocking
' workbook structure protection.
'
' @param enable True to allow delete; False to lock back.
'==================================================
Public Sub SetSheetDeleteEnabled(ByVal enable As Boolean)
    On Error GoTo EH

    If enable Then
        If ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Unprotect Password:=SYS_STRUCT_PWD
        End If
        MsgBox "Sheet delete is ENABLED. You can now delete sheets via Excel menu.", vbInformation
    Else
        If Not ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Protect Password:=SYS_STRUCT_PWD, Structure:=True
        End If
        MsgBox "Sheet delete is DISABLED (structure protected).", vbInformation
    End If

    Exit Sub

EH:
    MsgBox "SetSheetDeleteEnabled failed: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub EnableSheetDelete()
    SetSheetDeleteEnabled True
End Sub

Public Sub DisableSheetDelete()
    SetSheetDeleteEnabled False
End Sub
'run this to enable delete option for sheet
Public Sub EnableSheetDeleteWithPassword()
    Dim pw As String
    pw = InputBox("Enter admin password to enable sheet deletion:", "Admin")

    If pw = "" Then Exit Sub
    If pw <> SYS_STRUCT_PWD Then
        MsgBox "Wrong password.", vbCritical
        Exit Sub
    End If

    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect Password:=SYS_STRUCT_PWD
    End If

    MsgBox "Sheet delete ENABLED. Delete sheets via Excel menu now.", vbInformation
End Sub


