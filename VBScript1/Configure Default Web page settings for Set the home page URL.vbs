HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "Start Page"
strValue = "http://www.microsoft.com"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


