On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Download"
ValueName = "CheckExeSignatures"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Verify the identity of programs you download:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Verify the identity of programs you download: ", strValue
End If

