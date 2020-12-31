On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Main"
ValueName = "Disable Script Debugger"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Turn off the Script Debugger:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Turn off the Script Debugger: ", strValue
End If

