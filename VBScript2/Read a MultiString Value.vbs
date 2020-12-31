' Description: Uses WMI to read a multi-string registry value.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SYSTEM\CurrentControlSet\Services\Eventlog\System"
strValueName = "Sources"
oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath, _
    strValueName,arrValues
 
For Each strValue In arrValues
    Wscript.Echo  strValue
Next

