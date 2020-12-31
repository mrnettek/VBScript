Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Testsheet.xls")
objExcel.Visible = TRUE

Set objWorksheet = objWorkbook.Worksheets("Sheet1")
objWorksheet.Copy
  


