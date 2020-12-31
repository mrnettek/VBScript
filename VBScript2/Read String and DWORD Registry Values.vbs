' Description: Uses WMI to read a string and a DWORD registry value.


Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Console"
strValueName = "HistoryBufferSize"
oReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
Wscript.Echo "Current History Buffer Size: " & dwValue 
 
 
strKeyPath = "SOFTWARE\Microsoft\Windows Script Host\Settings"
strValueName = "TrustPolicy"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
Wscript.Echo "Current WSH Trust Policy Value: " & strValue

