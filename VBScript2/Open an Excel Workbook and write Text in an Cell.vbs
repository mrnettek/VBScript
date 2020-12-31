Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Temp\Test.xls")

Set objWorksheet = objWorkbook.Worksheets(1)
objWorksheet.Cells(3, 1).Value = "Test"


