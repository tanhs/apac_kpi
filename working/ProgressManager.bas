Attribute VB_Name = "ProgressManager"
Option Explicit

Private gProgress As ufProgress
Private gTotal As Long
Private gDone As Long

Public Sub ProgressStart(ByVal title As String, ByVal totalUnits As Long)
    gTotal = IIf(totalUnits <= 0, 1, totalUnits)
    gDone = 0

    Set gProgress = New ufProgress
    gProgress.InitUI title
    gProgress.Show vbModeless
    gProgress.SetProgress 0, "Starting..."
    
End Sub

Public Sub ProgressStep(ByVal stepText As String, Optional ByVal unitsDone As Long = 1)
    If gTotal <= 0 Then gTotal = 1
    gDone = gDone + unitsDone
    If gDone > gTotal Then gDone = gTotal

    Dim pct As Double
    pct = (gDone / gTotal) * 100#

    If Not gProgress Is Nothing Then
        gProgress.SetProgress pct, stepText
    End If
End Sub

Public Sub ProgressSet(ByVal pct As Double, ByVal stepText As String)
    If Not gProgress Is Nothing Then
        gProgress.SetProgress pct, stepText
    End If
End Sub

' If you use a userform named frmProgress, keep it consistent.
' If your form name differs, change it here.F
Public Sub ProgressForceClose()
    On Error Resume Next
    ' Hide first (avoid repaint issues)
    gProgress.Hide
    Unload gProgress
    On Error GoTo 0
End Sub

Public Sub ProgressFinish(ByVal msg As String)
    On Error Resume Next
    If Not gProgress Is Nothing Then
        'gProgress.lblStatus.caption = msg
        DoEvents
    End If
    ProgressForceClose
    On Error GoTo 0
End Sub
