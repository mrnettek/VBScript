Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objExcel.ScreenUpdating = False

For i = 1 to 100
    For j = 1 to 100
      objExcel.Cells(i,j) = i * j
    Next
Next
objExcel.ScreenUpdating = True
  


