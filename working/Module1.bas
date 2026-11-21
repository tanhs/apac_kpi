Attribute VB_Name = "Module1"

Option Explicit

Public Sub Debug_CheckRefs()
    Dim ws As Worksheet
    Dim n As name
    Dim c As Range
    Dim arr As Variant
    Dim i As Long

    Debug.Print "=== Workbook Names with #REF! ==="
    For Each n In ThisWorkbook.names
        If InStr(1, n.RefersTo, "#REF!", vbTextCompare) > 0 Then
            Debug.Print n.name, n.RefersTo
        End If
    Next n

    Debug.Print "=== Sheet Names with #REF! ==="
    For Each ws In ThisWorkbook.Worksheets
        For Each n In ws.names
            If InStr(1, n.RefersTo, "#REF!", vbTextCompare) > 0 Then
                Debug.Print ws.name & " -> " & n.name, n.RefersTo
            End If
        Next n
    Next ws

    Debug.Print "=== Formulas with #REF! ==="
    For Each ws In ThisWorkbook.Worksheets
        For Each c In ws.UsedRange.Cells
            If c.HasFormula Then
                If InStr(1, c.Formula, "#REF!", vbTextCompare) > 0 Then
                    Debug.Print ws.name & "!" & c.Address(0, 0), c.Formula
                End If
            End If
        Next c
    Next ws

    Debug.Print "=== Print_Area Names ==="
    For Each n In ThisWorkbook.names
        If InStr(1, n.name, "Print_Area", vbTextCompare) > 0 Then
            Debug.Print n.name, n.RefersTo
        End If
    Next n

    Debug.Print "=== External Links ==="
    arr = ThisWorkbook.LinkSources(xlExcelLinks)
    If IsEmpty(arr) Then
        Debug.Print "No external links"
    ElseIf IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            Debug.Print arr(i)
        Next i
    End If
End Sub

