On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
ValueName = "Use_Dlgbox_Colors"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Use Windows settings for Web page controls:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Use Windows settings for Web page controls: ", strValue
End If

