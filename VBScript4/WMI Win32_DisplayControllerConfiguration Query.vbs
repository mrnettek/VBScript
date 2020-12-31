On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DisplayControllerConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "BitsPerPixel: " & objItem.BitsPerPixel
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ColorPlanes: " & objItem.ColorPlanes
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceEntriesInAColorTable: " & objItem.DeviceEntriesInAColorTable
    Wscript.Echo "DeviceSpecificPens: " & objItem.DeviceSpecificPens
    Wscript.Echo "HorizontalResolution: " & objItem.HorizontalResolution
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "RefreshRate: " & objItem.RefreshRate
    Wscript.Echo "ReservedSystemPaletteEntries: " & objItem.ReservedSystemPaletteEntries
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SystemPaletteEntries: " & objItem.SystemPaletteEntries
    Wscript.Echo "VerticalResolution: " & objItem.VerticalResolution
    Wscript.Echo "VideoMode: " & objItem.VideoMode
Next

