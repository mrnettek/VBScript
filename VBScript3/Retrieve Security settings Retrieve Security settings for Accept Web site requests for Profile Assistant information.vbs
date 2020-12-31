On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Internet Explorer\Security\P3Global"
ValueName = "Enabled"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Accept Web site requests for Profile Assistant information:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Accept Web site requests for Profile Assistant information: ", dwValue
End If

