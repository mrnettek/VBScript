On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "ScreenSaveActive"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Verify whether or not a screen saver is being used:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Verify whether or not a screen saver is being used: ", strValue
End If

