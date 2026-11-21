Attribute VB_Name = "modFileUtil"
'========================================================
' Check if a file is locked/open by another process
' Returns TRUE if file is currently open (cannot write)
'========================================================
Public Function IsFileLocked(ByVal filePath As String) As Boolean

    Dim f As Integer
    f = FreeFile

    On Error Resume Next
    Open filePath For Binary Access Write Lock Read Write As #f
    If Err.Number <> 0 Then
        IsFileLocked = True
        Err.Clear
    Else
        Close #f
        IsFileLocked = False
    End If
    On Error GoTo 0

End Function
