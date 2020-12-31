Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

For Each objWorksheet in objWorkbook.Worksheets
    Wscript.Echo objWorksheet.Name
Next
  


