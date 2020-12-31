Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

Set objRange = objWorksheet.UsedRange

objRange.Replace "C:\Test\Image.jpg", "C:\Backup\Image.jpg"
  


