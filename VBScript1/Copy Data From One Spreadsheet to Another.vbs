Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("c:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets("Sheet1")

Set objRange = objWorksheet.Range("A1:A20")
objRange.Copy

Set objExcel2 = CreateObject("Excel.Application")
objExcel2.Visible = True

Set objWorkbook2 = objExcel2.Workbooks.Add
Set objWorksheet2 = objWorkbook2.Worksheets("Sheet1")

objWorksheet2.Paste
  


