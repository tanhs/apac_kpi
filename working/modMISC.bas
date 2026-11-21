Attribute VB_Name = "modMISC"
Option Explicit
'Things to do
'check remove / hide not required worksheets

'Release notes
'v25.9
'   - fixed Latest Update
'   - fixed Monthly Adjustment, moving whole month to AdjLog and insert 1 line adjustment
'   - ADD: delete scratch sheets (Pivots_Scratch, ChScratch_Tmp, Chart1)
'   - add workbook open, to refresh all slicer, pivots and charts
'V25.7
'   - fixed overall monthly and country chart not responding
'   - fixed country ach%
'V25.5
'   - fixed Monthly Budget vs Actual (Overall)secondary achivement% wrong plot
'   - to split country chart to bar and achivement %
'   - fix regerenate pdf
'V25.4
'   - fixed Monthly Budget vs Actual (Overall)secondary axis
'   - fixed Country Achievement (Overall) slicer Country & Month link to pivot
'   - fixed pdf country's Monthly Budget vs Actual (Overall) included
'V25.3
'   - fixed remove grand total from chart
'V25.2
'   - fix export pdf chart
'V25.1
'   - fix pivot month filter as mmm-yy
'   - fix chart output with non blank region
'V25
'   - fix region dropdown to country


'=========================
' BUTTON CREATION (RUN ONCE)
'=========================
Public Sub CreateBtn_ImportPrevMonthAdj()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    Dim shp As Shape
    On Error Resume Next
    ws.Shapes("btnImportPrevMonthAdj").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 620, 90, 180, 40)
    shp.name = "btnImportPrevMonthAdj"
    shp.TextFrame2.TextRange.Text = "Import Prev Month Adj"
    shp.TextFrame2.TextRange.Font.size = 12
    shp.Fill.ForeColor.RGB = RGB(16, 185, 129)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "Import_PrevMonthAdj"
End Sub
