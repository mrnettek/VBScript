Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\Domains\fabrikam.com"
objReg.CreateKey HKEY_CURRENT_USER,strKeyPath

strValueName = "*"
dwValue = 4
objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
  


