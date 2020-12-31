Dim arrExcelValues()

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
objExcel.Visible = True

i = 1
x = 0

Do Until objExcel.Cells(i, 1).Value = ""
    ReDim Preserve arrExcelValues(x)
    arrExcelValues(x) = objExcel.Cells(i, 1).Value
    i = i + 1
    x = x + 1
 Loop

objExcel.Quit

For Each strItem in arrExcelValues
    Wscript.Echo strItem
Next
  


