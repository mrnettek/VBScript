' Description: Demonstration script that adds the words "Test Value" to cell 1,1 in a new spreadsheet.


Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.Cells(1, 1).Value = "Test value"

