On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "AutoEndTasks"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Specify whether processes end when a user logs out:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Specify whether processes end when a user logs out: ", dwValue
End If

