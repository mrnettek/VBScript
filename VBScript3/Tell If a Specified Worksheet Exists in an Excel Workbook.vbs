Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

x = 0

For Each objWorksheet in objWorkbook.Worksheets
    If objWorksheet.Name = "Budget" Then
        x = 1
        Exit For
    End If
Next

objExcel.Quit

If x = 1 Then
    Wscript.Echo "The specified worksheet was found."
Else
    Wscript.Echo "The specified worksheet was not found."
End If
  


