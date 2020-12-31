Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Script Center"
objRegistry.CreateKey HKEY_CURRENT_USER, strKeyPath
  


