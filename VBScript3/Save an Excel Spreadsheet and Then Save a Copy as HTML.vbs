Const xlHTML = 44 

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objExcel.DisplayAlerts = False

objExcel.Cells(1, 1).Value = "A"
objExcel.Cells(1, 2).Value = "B"
objExcel.Cells(1, 3).Value = "C"
objExcel.Cells(1, 4).Value = "D"

objWorkbook.SaveAs "C:\Scripts\Test.xls"
objWorkbook.SaveAs "C:\Scripts\Test.htm", xlHTML
  


