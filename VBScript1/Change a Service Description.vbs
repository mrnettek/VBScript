Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "System\CurrentControlSet\Services\SerialKeys"
strValueName = "Description"
strValue = "This is the SerialKeys service."

objRegistry.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
  


