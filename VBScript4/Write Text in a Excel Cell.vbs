Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add

Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1, 1).Value = "Test"
  


