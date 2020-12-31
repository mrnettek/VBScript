On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer"
ValueName = "CaseSensitive"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Enable case-sensitive searches:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Enable case-sensitive searches: ", dwValue
End If

