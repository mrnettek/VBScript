HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "1803"
dwValue = 0
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue


