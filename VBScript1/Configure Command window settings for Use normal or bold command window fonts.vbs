HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Console"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "FontWeight"
dwValue = &H2bc
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue


