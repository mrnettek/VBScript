On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Console"
ValueName = "FontSize"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, ValueName, dwValue
If IsNull(strValue) Then
    Wscript.Echo "Set the size of the command window font:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Set the size of the command window font: ", dwValue
End If

