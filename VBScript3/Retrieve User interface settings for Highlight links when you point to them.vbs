On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Settings"
ValueName = "Use Anchor Hover Color"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Highlight links when you point to them:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Highlight links when you point to them: ", strValue
End If

