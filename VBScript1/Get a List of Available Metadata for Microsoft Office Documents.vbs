Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")

For Each strProperty in objWorkbook.BuiltInDocumentProperties
    Wscript.Echo strProperty.Name
Next
  


