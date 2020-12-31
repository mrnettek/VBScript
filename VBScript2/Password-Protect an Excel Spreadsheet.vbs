Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = FALSE

Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1, 1).Value = Now
objWorkbook.SaveAs "C:\Scripts\Test.xls",,"%reTG54w"
objExcel.Quit
  


