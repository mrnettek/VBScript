Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

x = 1

Do Until objExcel.Cells(x,1).Value = ""
    If CDate(objExcel.Cells(x,1).Value) < Date Then
        objExcel.Cells(x,1).Interior.ColorIndex = 3
    End If
    x = x + 1
 Loop
  


