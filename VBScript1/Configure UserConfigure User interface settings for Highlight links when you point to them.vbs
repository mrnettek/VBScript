HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Settings"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "Use Anchor Hover Color"
strValue = "yes"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


