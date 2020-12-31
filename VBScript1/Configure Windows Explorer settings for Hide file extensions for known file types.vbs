HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "HideFileExt"
dwValue = 1
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue


