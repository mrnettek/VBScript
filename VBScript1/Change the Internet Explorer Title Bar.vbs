Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
 
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
strValueName = "Window Title"
strValue = "The Scripting Guys"

objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strValue
  


