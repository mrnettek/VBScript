' Description: Uses WMI to read an expanded string registry value.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon"
strValueName = "UIHost"
oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strKeyPath, _
    strValueName,strValue
 
Wscript.Echo  "The Windows logon UI host is: " & strValue

