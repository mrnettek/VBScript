Const xlShiftToRight = -4161

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1,1) = "Dataset 1"
objWorksheet.Cells(1,2) = "Dataset 2"
objWorksheet.Cells(1,3) = "Dataset 4"

Set objRange = objExcel.Range("C1").EntireColumn
objRange.Insert(xlShiftToRight)
  


