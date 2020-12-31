On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Sound"
ValueName = "Beep"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Play a sound when an error occurs:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Play a sound when an error occurs: ", strValue
End If

