On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
ValueName = "Enable Browser Extensions"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Enable non-Microsoft browser extensions:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Enable non-Microsoft browser extensions: ", strValue
End If

