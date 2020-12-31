HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Ftp"
objReg.CreateKey HKEY_CURRENT_USER, strKeyPath
ValueName = "Use Web Based FTP"
strValue = "yes"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue


