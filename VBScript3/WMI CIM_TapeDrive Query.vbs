On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_TapeDrive",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Capabilities: " & objItem.Capabilities
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CompressionMethod: " & objItem.CompressionMethod
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "DefaultBlockSize: " & objItem.DefaultBlockSize
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "EOTWarningZoneSize: " & objItem.EOTWarningZoneSize
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxBlockSize: " & objItem.MaxBlockSize
    Wscript.Echo "MaxMediaSize: " & objItem.MaxMediaSize
    Wscript.Echo "MaxPartitionCount: " & objItem.MaxPartitionCount
    Wscript.Echo "MinBlockSize: " & objItem.MinBlockSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NeedsCleaning: " & objItem.NeedsCleaning
    Wscript.Echo "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
    Wscript.Echo "Padding: " & objItem.Padding
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

