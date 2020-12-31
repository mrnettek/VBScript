On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_PCVideoController",,48)
For Each objItem in colItems
    Wscript.Echo "AcceleratorCapabilities: " & objItem.AcceleratorCapabilities
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
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
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxMemorySupported: " & objItem.MaxMemorySupported
    Wscript.Echo "MaxNumberControlled: " & objItem.MaxNumberControlled
    Wscript.Echo "MaxRefreshRate: " & objItem.MaxRefreshRate
    Wscript.Echo "MinRefreshRate: " & objItem.MinRefreshRate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfColorPlanes: " & objItem.NumberOfColorPlanes
    Wscript.Echo "NumberOfVideoPages: " & objItem.NumberOfVideoPages
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ProtocolSupported: " & objItem.ProtocolSupported
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
    Wscript.Echo "VideoArchitecture: " & objItem.VideoArchitecture
    Wscript.Echo "VideoMemoryType: " & objItem.VideoMemoryType
    Wscript.Echo "VideoMode: " & objItem.VideoMode
    Wscript.Echo "VideoProcessor: " & objItem.VideoProcessor
Next

