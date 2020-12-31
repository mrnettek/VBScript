Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add

Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1, 1) = "A"
objWorksheet.Cells(1, 2) = "B"
objWorksheet.Cells(1, 3) = "C"

objWorksheet.Cells(1, 2).HorizontalAlignment = -4108
  


