Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
 
Set objRegistry=GetObject("winmgmts:\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Software\Test"
strValueName = "Sample Value 1"

objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath, strValueName
  


