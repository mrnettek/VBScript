On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
    Wscript.Echo "AddressWidth: " & objItem.AddressWidth
    Wscript.Echo "Architecture: " & objItem.Architecture
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CpuStatus: " & objItem.CpuStatus
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentClockSpeed: " & objItem.CurrentClockSpeed
    Wscript.Echo "CurrentVoltage: " & objItem.CurrentVoltage
    Wscript.Echo "DataWidth: " & objItem.DataWidth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ExtClock: " & objItem.ExtClock
    Wscript.Echo "Family: " & objItem.Family
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "L2CacheSize: " & objItem.L2CacheSize
    Wscript.Echo "L2CacheSpeed: " & objItem.L2CacheSpeed
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Level: " & objItem.Level
    Wscript.Echo "LoadPercentage: " & objItem.LoadPercentage
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxClockSpeed: " & objItem.MaxClockSpeed
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherFamilyDescription: " & objItem.OtherFamilyDescription
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ProcessorId: " & objItem.ProcessorId
    Wscript.Echo "ProcessorType: " & objItem.ProcessorType
    Wscript.Echo "Revision: " & objItem.Revision
    Wscript.Echo "Role: " & objItem.Role
    Wscript.Echo "SocketDesignation: " & objItem.SocketDesignation
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "Stepping: " & objItem.Stepping
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "UniqueId: " & objItem.UniqueId
    Wscript.Echo "UpgradeMethod: " & objItem.UpgradeMethod
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "VoltageCaps: " & objItem.VoltageCaps
Next

