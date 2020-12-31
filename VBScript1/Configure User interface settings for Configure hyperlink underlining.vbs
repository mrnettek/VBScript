HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "Anchor Underline"
strValue = "yes"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


