Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set colSheets = objWorkbook.Sheets

colSheets.Add ,,9
  


