Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

For i = 1 to 10
    objExcel.Cells(i, 1).Value = i/6
Next

Set objRange = objWorksheet.UsedRange
objRange.NumberFormat = "#.0000"
  


