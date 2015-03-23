Sub CATMain()
On Error Resume Next
Set drawingDocument1 = CATIA.ActiveDocument
Set drawingSheets1 = drawingDocument1.Sheets
Set drawingSheet1 = drawingSheets1.ActiveSheet
Set drawingViews1 = drawingSheet1.Views
Set ActView = drawingViews1.ActiveView
Set ActTables = ActView.Tables
Dim m As Integer
m = 1
msgboxtext = "Warning!"
MsgBox "If you want to copy the values"& vbCrLf &"first you have to activate the view where they are." , , msgboxtext

'-----------Open Excel-----------------------------------------------
Set excell = CreateObject("Excel.Application")
excell.Visible = True
Set excelWorkbooks= excell.Workbooks.Add

'-----------Loop 1 | All tables in 2D----------------------
For i = 1 To ActTables.Count
    Set drawingTable1 = ActTables.Item(i)

    Dim int1 As Integer
    
    int1 = drawingTable1.NumberOfColumns

    Dim int2 As Integer
    
    int2 = drawingTable1.NumberOfRows

'-----------Loop 2 | Values to cells---------------------
    For Row = 1 To long3
        For Col = 1 To long2
            excell.Cells(m, Col).Value = drawingTable1.GetCellString(Row, Col)
        Next Col
    m = m + 1
    Next Row
Next i
End Sub
