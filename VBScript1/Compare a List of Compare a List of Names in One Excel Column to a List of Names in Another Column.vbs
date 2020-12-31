Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

Set objRange = objWorksheet.Range("B1").EntireColumn
i = 1

Do Until objExcel.Cells(i, 1).Value = ""
    strName = objExcel.Cells(i, 1).Value
    Set objSearch = objRange.Find(strName)

    If objSearch Is Nothing Then
        Wscript.Echo strName & " was not found."
    Else
        Wscript.Echo strName & " was found."
    End If

    i = i + 1
Loop
  


