Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.UsedRange.ClearContents
  


