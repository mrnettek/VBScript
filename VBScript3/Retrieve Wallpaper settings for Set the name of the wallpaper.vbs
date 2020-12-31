On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "Wallpaper"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Set the name of the wallpaper:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Set the name of the wallpaper: ", strValue
End If

