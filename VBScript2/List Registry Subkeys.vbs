' Description: Uses WMI to enumerate all the registry subkeys under HKLM\SYSTEM\CurrentControlSet\Services.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SYSTEM\CurrentControlSet\Services"
oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
 
For Each subkey In arrSubKeys
    Wscript.Echo subkey
Next

