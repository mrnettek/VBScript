Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

For x = 1 to 10
    For y = 1 to 10
        objExcel.Cells(x, y).Value = x + y
    Next
Next

objWorksheet.PageSetup.PrintArea = "B2:D4"
  


