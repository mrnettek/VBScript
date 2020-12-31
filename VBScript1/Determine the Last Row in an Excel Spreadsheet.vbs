Const xlCellTypeLastCell = 11

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)
objWorksheet.Activate

Set objRange = objWorksheet.UsedRange
objRange.SpecialCells(xlCellTypeLastCell).Activate

intNewRow = objExcel.ActiveCell.Row + 1
strNewCell = "A" &  intNewRow

objExcel.Range(strNewCell).Activate
  


