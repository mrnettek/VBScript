Const HKEY_LOCAL_MACHINE = &H80000002
 
strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC DATA SOURCES"

objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
 
For i = 0 to Ubound(arrValueNames)
    strValueName = arrValueNames(i)
    objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue    
    Wscript.Echo strValueName & " -- " & strValue
Next
  


