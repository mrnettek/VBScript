Set objExcel = CreateObject("Excel.Application") 

objExcel.Visible = TRUE 
objExcel.DisplayAlerts = FALSE

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls",,,,"L$6tg4HHE")

objWorkbook.Password = ""
objWorkbook.SaveAs "C:\Scripts\Test.xls"
  


