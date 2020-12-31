Const xlShiftDown = -4121

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

i = 1

strStartValue = Left(objExcel.Cells(i, 1), 1)

Do Until objExcel.Cells(i, 1) = ""
    strValue = Left(objExcel.Cells(i, 1), 1)
    If strValue <> strStartValue Then
        Set objRange = objExcel.Cells(i,1).EntireRow
        objRange.Activate
        objRange.Insert xlShiftDown
        strStartValue = Left(objExcel.Cells(i + 1, 1), 1)
    End If
    i = i + 1
Loop
  


