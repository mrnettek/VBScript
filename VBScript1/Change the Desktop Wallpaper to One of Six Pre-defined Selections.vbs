Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

intLowNumber = 1
intHighNumber = 6

Randomize

intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)

Select Case intNumber
    Case 1
        strValue = "C:\WINDOWS\System32\Wallpaper1.bmp"
    Case 2
        strValue = "C:\WINDOWS\System32\Wallpaper2.bmp"
    Case 3
        strValue = "C:\WINDOWS\System32\Wallpaper3.bmp"
    Case 4
        strValue = "C:\WINDOWS\System32\Wallpaper4.bmp"
    Case 5
        strValue = "C:\WINDOWS\System32\Wallpaper5.bmp"
    Case 6
        strValue = "C:\WINDOWS\System32\Wallpaper6.bmp"
End Select

strKeyPath = "Control Panel\Desktop"
ValueName = "Wallpaper"

objReg.SetStringValue HKEY_USERS, strKeyPath, ValueName, strValue
  


