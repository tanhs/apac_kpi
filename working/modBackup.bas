Attribute VB_Name = "modBackup"
'modBackup module
Option Explicit

'==================================================
' Backup helpers
'==================================================
Private Const OD_LIBRARY_ROOT As String = "Malaysia Information Technology - APAC KPI"




'--------------------------------------------------
' Generates token like: 2026_02_21_193833_797
' (ms-ish + a little randomness to avoid collisions)
'--------------------------------------------------
Public Function backupToken() As String
    Dim ms As Long
    ms = (Timer - Fix(Timer)) * 1000

    Dim rnd3 As Long
    Randomize
    rnd3 = CLng(Rnd() * 999)

    backupToken = Format$(Now, "yyyy_mm_dd_hhnnss") & "_" & _
                  Format$(ms Mod 1000, "000") & "_" & _
                  Format$(rnd3 Mod 1000, "000")
End Function

'--------------------------------------------------
' Makes a SaveCopyAs into /backup
'
' @param stageSuffix "Before" or "After"
' @param token Pass same token for both backups in one run
'--------------------------------------------------
Public Sub BackupWorkbookCopy(ByVal stageSuffix As String, ByVal token As String)
    Dim folderPath As String
    folderPath = EnsureBackupFolderPath()

    Dim baseName As String
    baseName = ThisWorkbook.name
    baseName = Replace(baseName, ".xlsm", "", , , vbTextCompare)
    baseName = Replace(baseName, ".xlsx", "", , , vbTextCompare)
    baseName = Replace(baseName, ".xls", "", , , vbTextCompare)

    Dim outFile As String
    outFile = folderPath & "\" & baseName & "_" & token & "_" & stageSuffix & ".xlsm"

    ThisWorkbook.SaveCopyAs outFile
End Sub

Public Sub EnsureFolderRecursive(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim cur As String

    folderPath = Replace(folderPath, "/", "\")
    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then Exit Sub

    parts = Split(folderPath, "\")
    cur = parts(0)

    ' Handle "C:" root
    If Right$(cur, 1) = ":" Then cur = cur & "\"

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Right$(cur, 1) <> "\" Then cur = cur & "\"
            cur = cur & parts(i)

            On Error Resume Next
            If Dir$(cur, vbDirectory) = vbNullString Then MkDir cur
            On Error GoTo 0
        End If
    Next i
End Sub

'--------------------------------------------------
' Returns local backup folder:
'   <OneDriveCommercial>\<OD_LIBRARY_ROOT>\backup
' Works even when workbook is opened from SharePoint URL.
'--------------------------------------------------
Public Function EnsureBackupFolderPath() As String
    Dim oneDriveRoot As String
    oneDriveRoot = Environ$("OneDriveCommercial")
    If Len(oneDriveRoot) = 0 Then oneDriveRoot = Environ$("OneDrive")

    If Len(oneDriveRoot) = 0 Then
        Err.Raise vbObjectError + 8701, , _
            "OneDrive is not available on this PC. Cannot create local backups."
    End If

    Dim baseFolder As String
    baseFolder = oneDriveRoot & "\" & OD_LIBRARY_ROOT

    Dim backupPath As String
    backupPath = baseFolder & "\backup"

    ' Create all levels safely
    EnsureFolderRecursive backupPath

    ' Validate it exists now
    On Error Resume Next
    Dim ok As Boolean
    ok = (Dir$(backupPath, vbDirectory) <> vbNullString)
    On Error GoTo 0

    If Not ok Then
        Err.Raise vbObjectError + 8702, , _
            "Backup folder could not be created or accessed:" & vbCrLf & backupPath
    End If

    EnsureBackupFolderPath = backupPath
End Function

