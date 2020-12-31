Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

For i = 1 to 5
    If i Mod 2 = 0 Then
        Set objRange = objExcel.ActiveCell.EntireRow
        objRange.Cells.Interior.ColorIndex = 37
    Else
        Set objRange = objExcel.ActiveCell.EntireRow
        objRange.Cells.Interior.ColorIndex = 36
    End If
        
    objWorksheet.Cells(i,1) = i

    intNewRow = objExcel.ActiveCell.Row + 1
    strNewCell = "A" &  intNewRow
    objExcel.Range(strNewCell).Activate

Next
  


