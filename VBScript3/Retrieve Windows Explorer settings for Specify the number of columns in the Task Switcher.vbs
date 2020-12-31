On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "CoolSwitchColumns"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Specify the number of columns in the Task Switcher:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Specify the number of columns in the Task Switcher: ", strValue
End If

