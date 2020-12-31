On Error Resume Next
HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Control Panel\Mouse"
ValueName = "SwapMouseButtons"
    objReg.GetStringValue HKEY_CURRENT_USER, strKeyPath, ValueName, strValue
If IsNull(strValue) Then
    Wscript.Echo "Swap the left and right mouse buttons:  The value is either Null or could not be found in the registry."
Else
    Wscript.Echo "Swap the left and right mouse buttons: ", strValue
End If

