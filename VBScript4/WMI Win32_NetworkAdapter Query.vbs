On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)
For Each objItem in colItems
    Wscript.Echo "AdapterType: " & objItem.AdapterType
    Wscript.Echo "AdapterTypeId: " & objItem.AdapterTypeId
    Wscript.Echo "AutoSense: " & objItem.AutoSense
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "Index: " & objItem.Index
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Installed: " & objItem.Installed
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MACAddress: " & objItem.MACAddress
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxNumberControlled: " & objItem.MaxNumberControlled
    Wscript.Echo "MaxSpeed: " & objItem.MaxSpeed
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NetConnectionID: " & objItem.NetConnectionID
    Wscript.Echo "NetConnectionStatus: " & objItem.NetConnectionStatus
    Wscript.Echo "NetworkAddresses: " & objItem.NetworkAddresses
    Wscript.Echo "PermanentAddress: " & objItem.PermanentAddress
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ProductName: " & objItem.ProductName
    Wscript.Echo "ServiceName: " & objItem.ServiceName
    Wscript.Echo "Speed: " & objItem.Speed
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
Next

