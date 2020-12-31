On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "FontSmoothingType"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Specify ClearType or standard font smoothing:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Specify ClearType or standard font smoothing: ", dwValue
End If

