On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Desktop",,48)
For Each objItem in colItems
    Wscript.Echo "BorderWidth: " & objItem.BorderWidth
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CoolSwitch: " & objItem.CoolSwitch
    Wscript.Echo "CursorBlinkRate: " & objItem.CursorBlinkRate
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DragFullWindows: " & objItem.DragFullWindows
    Wscript.Echo "GridGranularity: " & objItem.GridGranularity
    Wscript.Echo "IconSpacing: " & objItem.IconSpacing
    Wscript.Echo "IconTitleFaceName: " & objItem.IconTitleFaceName
    Wscript.Echo "IconTitleSize: " & objItem.IconTitleSize
    Wscript.Echo "IconTitleWrap: " & objItem.IconTitleWrap
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Pattern: " & objItem.Pattern
    Wscript.Echo "ScreenSaverActive: " & objItem.ScreenSaverActive
    Wscript.Echo "ScreenSaverExecutable: " & objItem.ScreenSaverExecutable
    Wscript.Echo "ScreenSaverSecure: " & objItem.ScreenSaverSecure
    Wscript.Echo "ScreenSaverTimeout: " & objItem.ScreenSaverTimeout
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "Wallpaper: " & objItem.Wallpaper
    Wscript.Echo "WallpaperStretched: " & objItem.WallpaperStretched
    Wscript.Echo "WallpaperTiled: " & objItem.WallpaperTiled
Next

