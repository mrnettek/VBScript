Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Temp\Test.xls")

For Each objWorksheet in objWorkbook.Worksheets
    	If objWorksheet.Name = "Sheet2" then
		objWorksheet.Activate
		objWorksheet.Cells(3, 1).Value = "Test"
	End if
Next
  


