' Description: Uses WMI to check access rights for the logged on user to the HKLM\SYSTEM\CurrentControlSet portion of the registry.


Const KEY_QUERY_VALUE = &H0001
Const KEY_SET_VALUE = &H0002
Const KEY_CREATE_SUB_KEY = &H0004
Const DELETE = &H00010000
 
Const HKEY_LOCAL_MACHINE = &H80000002
 
strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SYSTEM\CurrentControlSet"
 
oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_QUERY_VALUE, _
    bHasAccessRight
If bHasAccessRight = True Then
    Wscript.Echo "Have Query Value Access Rights on Key"
Else
    Wscript.Echo "Do Not Have Query Value Access Rights on Key"
End If   
 
oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_SET_VALUE, _
    bHasAccessRight
If bHasAccessRight = True Then
    Wscript.Echo "Have Set Value Access Rights on Key"
Else
    Wscript.Echo "Do Not Have Set Value Access Rights on Key"
End If   
 
oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, KEY_CREATE_SUB_KEY, _
    bHasAccessRight
If bHasAccessRight = True Then
    Wscript.Echo "Have Create SubKey Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Create SubKey Access Rights on Key"
End If
 
oReg.CheckAccess HKEY_LOCAL_MACHINE, strKeyPath, DELETE, bHasAccessRight
If bHasAccessRight = True Then
    StdOut.WriteLine "Have Delete Access Rights on Key"
Else
    StdOut.WriteLine "Do Not Have Delete Access Rights on Key"
End If

