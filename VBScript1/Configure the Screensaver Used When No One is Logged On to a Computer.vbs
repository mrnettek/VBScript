Const HKEY_USERS = &H80000003

strComputer = "."

Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = ".Default\Control Panel\Desktop"
ValueName = "SCRNSAVE.EXE"
strValue = "C:\WINDOWS\System32\Script Center.scr"

objReg.SetStringValue HKEY_USERS, strKeyPath, ValueName, strValue
  


