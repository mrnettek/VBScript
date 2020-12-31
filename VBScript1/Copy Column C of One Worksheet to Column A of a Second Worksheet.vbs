Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

Set objWorksheet = objWorkbook.Worksheets(2)
objWorksheet.Activate

Set objRange = objWorkSheet.Range("C1").EntireColumn
objRange.Copy

Set objWorksheet = objWorkbook.Worksheets(1)
objWorksheet.Activate

Set objRange = objWorkSheet.Range("A1")
objWorksheet.Paste(objRange)
  


