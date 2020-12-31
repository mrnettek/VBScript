HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "ProxyHttp1.1"
dwValue = 1
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue


