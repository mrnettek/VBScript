HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "FontSmoothingType"
dwValue = 1
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue


