' Description: Demonstration script that opens an existing Excel spreadsheet named C:\Scripts\New_users.xls.


Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\New_users.xls")

