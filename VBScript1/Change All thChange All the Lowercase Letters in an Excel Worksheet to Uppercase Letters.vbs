Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1,1) = "abcdef"
objWorksheet.Cells(1,2) = "ghijkl"
objWorksheet.Cells(1,3) = "mnopqr"
objWorksheet.Cells(1,4) = "stuvwx"

Wscript.Sleep 2000

Set objRange = objWorksheet.UsedRange

For Each objCell in objRange
    objCell.Value = UCase(objCell.Value)
Next
  


