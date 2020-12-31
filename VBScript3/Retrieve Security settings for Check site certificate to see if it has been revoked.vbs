On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
ValueName = "CertificateRevocation"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Check site certificate to see if it has been revoked:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Check site certificate to see if it has been revoked: ", dwValue
End If

