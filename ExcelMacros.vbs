Sub AutoFitAllSheets()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate
ws.Cells.EntireColumn.AutoFit
Next ws
End Sub