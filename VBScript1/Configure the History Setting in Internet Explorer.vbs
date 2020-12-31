HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\URL History"
ValueName = "DaysToKeep"
dwValue = 25

objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
  


