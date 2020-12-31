Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

i = 1

Do Until objWorksheet.Cells(i, 1) = ""
    intA = objWorksheet.Cells(i, 1)
    intB = objWorksheet.Cells(i, 2)
    If intB - intA >= 10 Then
        objWorksheet.Cells(i, 1).Font.ColorIndex = 3
        objWorksheet.Cells(i, 2).Font.ColorIndex = 3
    End If
    i = i + 1
Loop
  


