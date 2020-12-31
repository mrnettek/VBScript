' Print a Excel Sheet
' ------------------------ Excel Output
Dim oExcel, oSheet


Call Init_XLS()

	oSheet.Cells(1,1).Value = "This is a Test Print"
	oSheet.PrintOut

Call Close_XLS()



'-------------
Sub Init_XLS()

	Set oExcel = CreateObject("Excel.application")
	oExcel.Visible = False
	oExcel.Workbooks.add

	Set oSheet = oExcel.ActiveWorkbook.Worksheets(1)
	oSheet.Name = "TestPage"

	oExcel.Worksheets(3).Delete
	oExcel.Worksheets(2).Delete

	oExcel.DisplayAlerts = False

End Sub


'--------------
Sub Close_XLS()

	oExcel.ActiveWorkbook.Close False
	oExcel.Quit

	set oSheet = Nothing
	Set oExcel = Nothing

End Sub

