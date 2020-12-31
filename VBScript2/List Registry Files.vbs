' Description: Uses WMI to list all the registry file and locations under HKLM\System\CurrentControlSet\Control\Hivelist.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
strKeyPath = "System\CurrentControlSet\Control\hivelist"
oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath,_
    arrValueNames, arrValueTypes
 
For i=0 To UBound(arrValueNames)
    StdOut.WriteLine "File Name: " & arrValueNames(i) & " -- "      
    oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath, _
        arrValueNames(i),strValue
    Wscript.Echo "Location: " & strValue
    Wscript.Echo
Next

