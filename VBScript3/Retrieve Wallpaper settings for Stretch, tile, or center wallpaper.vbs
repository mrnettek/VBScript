On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Desktop"
ValueName = "WallpaperStyle"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Stretch, tile, or center wallpaper:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Stretch, tile, or center wallpaper: ", strValue
End If

