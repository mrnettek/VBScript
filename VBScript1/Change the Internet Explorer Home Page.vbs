Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
ValueName = "Start Page"
strValue = "http://www.microsoft.com/technet/scriptcenter/default.mspx"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
  


