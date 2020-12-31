On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\FTP"
ValueName = "Use PASV"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Use passive FTP:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Use passive FTP: ", strValue
End If

