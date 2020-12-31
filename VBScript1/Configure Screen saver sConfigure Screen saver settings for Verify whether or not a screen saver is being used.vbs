HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "ScreenSaveActive"
strValue = "1"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


