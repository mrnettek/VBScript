Const xlAscending = 1
Const xlNo = 2
Const xlSortRows = 2

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1, 1) = "C"
objWorksheet.Cells(1, 2) = "D"
objWorksheet.Cells(1, 3) = "B"
objWorksheet.Cells(1, 4) = "A"

Set objRange = objExcel.ActiveCell.EntireRow
objRange.Sort objRange, xlAscending, , , , , , xlNo, , , xlSortRows
  


