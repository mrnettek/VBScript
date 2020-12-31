On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState"
ValueName = "FullPathAddress"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Show full folder path in the address bar:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Show full folder path in the address bar: ", dwValue
End If

