Const acExport = 1
Const acSpreadsheetTypeExcel9 = 8

Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase "C:\Scripts\Test.mdb"

objAccess.DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
    "Employees", "C:\Scripts\Employees.xls", True
  


