' Description: Uses WMI to list all the registry values and their types under HKLM\SYSTEM\CurrentControlSet\Control\Lsa.


Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
 
strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SYSTEM\CurrentControlSet\Control\Lsa"
 
oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, _
    arrValueNames, arrValueTypes
 
For i=0 To UBound(arrValueNames)
    Wscript.Echo "Value Name: " & arrValueNames(i) 
    
    Select Case arrValueTypes(i)
        Case REG_SZ
            Wscript.Echo "Data Type: String"
            Wscript.Echo
        Case REG_EXPAND_SZ
            Wscript.Echo "Data Type: Expanded String"
            Wscript.Echo
        Case REG_BINARY
            Wscript.Echo "Data Type: Binary"
            Wscript.Echo
        Case REG_DWORD
            Wscript.Echo "Data Type: DWORD"
            Wscript.Echo
        Case REG_MULTI_SZ
            Wscript.Echo "Data Type: Multi String"
            Wscript.Echo
    End Select 
Next

