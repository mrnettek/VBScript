Const xlCSV = 6

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Testsheet.xls")
objExcel.DisplayAlerts = FALSE
objExcel.Visible = TRUE

Set objWorksheet = objWorkbook.Worksheets("Sheet1")
objWorksheet.SaveAs "c:\scripts\test.csv", xlCSV

objExcel.Quit
  


