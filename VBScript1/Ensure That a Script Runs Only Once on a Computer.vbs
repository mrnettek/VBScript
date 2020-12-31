Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\AdminScripts"
objRegistry.CreateKey HKEY_LOCAL_MACHINE, strKeyPath

strValueName = "Script 1"
strValue = "No"
objRegistry.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
  


