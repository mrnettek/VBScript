On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Mouse"
ValueName = "MouseTrails"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Display mouse pointer trails:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Display mouse pointer trails: ", strValue
End If

