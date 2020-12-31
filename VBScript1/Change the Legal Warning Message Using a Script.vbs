Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
 
Set objReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
strValueName = "LegalNoticeCaption"
strValue = "Fabrikam, Inc. Legal Notice"
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
strValueName = "LegalNoticeText"
strValue = "By logging on to this computer you agree to abide by the "
strValue = strValue & "computer usage rules and regulations of Fabrikam, Inc."
strValue = strValue & vbCrLf & vbCrLf
strValue = strValue & "For more information, phone (425)-555-1289."
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
  


