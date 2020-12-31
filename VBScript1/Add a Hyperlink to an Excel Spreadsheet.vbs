Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1, 1) = "Script Center"

Set objRange = objExcel.Range("A1")
Set objLink = objWorksheet.Hyperlinks.Add _
    (objRange, "http://www.microsoft.com/technet/scriptcenter")
  


