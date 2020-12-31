Const acImport = 0
Const acSpreadsheetTypeExcel9 = 8

Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase "C:\Scripts\Test.mdb"

objAccess.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
    "Employees", "C:\Scripts\Employees.xls", True
  


