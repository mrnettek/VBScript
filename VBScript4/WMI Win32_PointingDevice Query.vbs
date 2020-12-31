On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PointingDevice",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DeviceInterface: " & objItem.DeviceInterface
    Wscript.Echo "DoubleSpeedThreshold: " & objItem.DoubleSpeedThreshold
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "Handedness: " & objItem.Handedness
    Wscript.Echo "HardwareType: " & objItem.HardwareType
    Wscript.Echo "InfFileName: " & objItem.InfFileName
    Wscript.Echo "InfSection: " & objItem.InfSection
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IsLocked: " & objItem.IsLocked
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfButtons: " & objItem.NumberOfButtons
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PointingType: " & objItem.PointingType
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "QuadSpeedThreshold: " & objItem.QuadSpeedThreshold
    Wscript.Echo "Resolution: " & objItem.Resolution
    Wscript.Echo "SampleRate: " & objItem.SampleRate
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "Synch: " & objItem.Synch
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

