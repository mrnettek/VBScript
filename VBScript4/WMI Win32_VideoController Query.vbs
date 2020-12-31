On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController",,48)
For Each objItem in colItems
    Wscript.Echo "AcceleratorCapabilities: " & objItem.AcceleratorCapabilities
    Wscript.Echo "AdapterCompatibility: " & objItem.AdapterCompatibility
    Wscript.Echo "AdapterDACType: " & objItem.AdapterDACType
    Wscript.Echo "AdapterRAM: " & objItem.AdapterRAM
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ColorTableEntries: " & objItem.ColorTableEntries
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentBitsPerPixel: " & objItem.CurrentBitsPerPixel
    Wscript.Echo "CurrentHorizontalResolution: " & objItem.CurrentHorizontalResolution
    Wscript.Echo "CurrentNumberOfColors: " & objItem.CurrentNumberOfColors
    Wscript.Echo "CurrentNumberOfColumns: " & objItem.CurrentNumberOfColumns
    Wscript.Echo "CurrentNumberOfRows: " & objItem.CurrentNumberOfRows
    Wscript.Echo "CurrentRefreshRate: " & objItem.CurrentRefreshRate
    Wscript.Echo "CurrentScanMode: " & objItem.CurrentScanMode
    Wscript.Echo "CurrentVerticalResolution: " & objItem.CurrentVerticalResolution
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DeviceSpecificPens: " & objItem.DeviceSpecificPens
    Wscript.Echo "DitherType: " & objItem.DitherType
    Wscript.Echo "DriverDate: " & objItem.DriverDate
    Wscript.Echo "DriverVersion: " & objItem.DriverVersion
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ICMIntent: " & objItem.ICMIntent
    Wscript.Echo "ICMMethod: " & objItem.ICMMethod
    Wscript.Echo "InfFilename: " & objItem.InfFilename
    Wscript.Echo "InfSection: " & objItem.InfSection
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InstalledDisplayDrivers: " & objItem.InstalledDisplayDrivers
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxMemorySupported: " & objItem.MaxMemorySupported
    Wscript.Echo "MaxNumberControlled: " & objItem.MaxNumberControlled
    Wscript.Echo "MaxRefreshRate: " & objItem.MaxRefreshRate
    Wscript.Echo "MinRefreshRate: " & objItem.MinRefreshRate
    Wscript.Echo "Monochrome: " & objItem.Monochrome
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfColorPlanes: " & objItem.NumberOfColorPlanes
    Wscript.Echo "NumberOfVideoPages: " & objItem.NumberOfVideoPages
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ProtocolSupported: " & objItem.ProtocolSupported
    Wscript.Echo "ReservedSystemPaletteEntries: " & objItem.ReservedSystemPaletteEntries
    Wscript.Echo "SpecificationVersion: " & objItem.SpecificationVersion
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "SystemPaletteEntries: " & objItem.SystemPaletteEntries
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
    Wscript.Echo "VideoArchitecture: " & objItem.VideoArchitecture
    Wscript.Echo "VideoMemoryType: " & objItem.VideoMemoryType
    Wscript.Echo "VideoMode: " & objItem.VideoMode
    Wscript.Echo "VideoModeDescription: " & objItem.VideoModeDescription
    Wscript.Echo "VideoProcessor: " & objItem.VideoProcessor
Next

