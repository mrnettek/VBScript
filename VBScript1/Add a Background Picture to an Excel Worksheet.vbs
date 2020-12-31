Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)
objWorksheet.SetBackgroundPicture "C:\Scripts\Test.jpg"
  


