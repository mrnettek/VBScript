HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Mouse"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "MouseTrails"
strValue = "4"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


