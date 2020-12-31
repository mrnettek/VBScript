Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon"
strValueName = "DefaultUserName"

objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue

Wscript.Echo strValue
  


