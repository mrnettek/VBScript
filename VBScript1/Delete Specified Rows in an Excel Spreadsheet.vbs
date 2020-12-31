Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

i = 1

Do Until objExcel.Cells(i, 1).Value = ""
    If objExcel.Cells(i, 1).Value = "delete" Then
        Set objRange = objExcel.Cells(i, 1).EntireRow
        objRange.Delete
        i = i - 1
    End If
    i = i + 1
Loop
  


