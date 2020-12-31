HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "AutoSuggest"
strValue = "yes"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


