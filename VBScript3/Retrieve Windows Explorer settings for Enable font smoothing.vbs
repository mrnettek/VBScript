On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "FontSmoothing"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Enable font smoothing:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Enable font smoothing: ", strValue
End If

