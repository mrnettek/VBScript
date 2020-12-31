On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Ftp"
ValueName = "Use Web Based FTP"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Show FTP sites in folder view:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Show FTP sites in folder view: ", strValue
End If

