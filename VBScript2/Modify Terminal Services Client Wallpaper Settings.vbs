' Description: Disables client wallpaper on a computer running Terminal Services. To enable the use of client wallpaper, pass the value 1 (rather than 0) to the SetClientWallpaper method.


Const NO_WALLPAPER = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSEnvironmentSetting")

For Each objItem in colItems
    errResult = objItem.SetClientWallpaper(NO_WALLPAPER)
Next

