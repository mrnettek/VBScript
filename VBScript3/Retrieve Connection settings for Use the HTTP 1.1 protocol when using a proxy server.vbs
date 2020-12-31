On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
ValueName = "ProxyHttp1.1"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Use the HTTP 1.1 protocol when using a proxy server:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Use the HTTP 1.1 protocol when using a proxy server: ", dwValue
End If

