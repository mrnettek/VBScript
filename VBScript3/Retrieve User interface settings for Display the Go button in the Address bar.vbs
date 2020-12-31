On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Internet Explorer\Main"
ValueName = "ShowGoButton"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Display the Go button in the Address bar:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Display the Go button in the Address bar: ", strValue
End If

