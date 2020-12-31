On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "ActualColorResolution: " & objItem.ActualColorResolution
    Wscript.Echo "AdapterChipType: " & objItem.AdapterChipType
    Wscript.Echo "AdapterCompatibility: " & objItem.AdapterCompatibility
    Wscript.Echo "AdapterDACType: " & objItem.AdapterDACType
    Wscript.Echo "AdapterDescription: " & objItem.AdapterDescription
    Wscript.Echo "AdapterRAM: " & objItem.AdapterRAM
    Wscript.Echo "AdapterType: " & objItem.AdapterType
    Wscript.Echo "BitsPerPixel: " & objItem.BitsPerPixel
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ColorPlanes: " & objItem.ColorPlanes
    Wscript.Echo "ColorTableEntries: " & objItem.ColorTableEntries
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceSpecificPens: " & objItem.DeviceSpecificPens
    Wscript.Echo "DriverDate: " & objItem.DriverDate
    Wscript.Echo "HorizontalResolution: " & objItem.HorizontalResolution
    Wscript.Echo "InfFilename: " & objItem.InfFilename
    Wscript.Echo "InfSection: " & objItem.InfSection
    Wscript.Echo "InstalledDisplayDrivers: " & objItem.InstalledDisplayDrivers
    Wscript.Echo "MonitorManufacturer: " & objItem.MonitorManufacturer
    Wscript.Echo "MonitorType: " & objItem.MonitorType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PixelsPerXLogicalInch: " & objItem.PixelsPerXLogicalInch
    Wscript.Echo "PixelsPerYLogicalInch: " & objItem.PixelsPerYLogicalInch
    Wscript.Echo "RefreshRate: " & objItem.RefreshRate
    Wscript.Echo "ScanMode: " & objItem.ScanMode
    Wscript.Echo "ScreenHeight: " & objItem.ScreenHeight
    Wscript.Echo "ScreenWidth: " & objItem.ScreenWidth
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SystemPaletteEntries: " & objItem.SystemPaletteEntries
    Wscript.Echo "VerticalResolution: " & objItem.VerticalResolution
Next

