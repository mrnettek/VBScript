Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

i = 1

Do Until objExcel.Cells(i, 1) = ""
    strValue = objExcel.Cells(i, 1)

    If IsDate(strValue) Then
        objExcel.Cells(i, 1).EntireRow.Interior.ColorIndex = 44
    End If
    
    i = i + 1
Loop
  


