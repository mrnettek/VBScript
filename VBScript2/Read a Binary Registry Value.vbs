' Description: Uses WMI to read a binary registry value.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set StdOut = WScript.StdOut
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "LicenseInfo"
oReg.GetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath, _
    strValueName,strValue
 
 
For i = lBound(strValue) to uBound(strValue)
    StdOut.WriteLine  strValue(i)
Next

