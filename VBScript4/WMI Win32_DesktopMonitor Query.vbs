On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Bandwidth: " & objItem.Bandwidth
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DisplayType: " & objItem.DisplayType
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IsLocked: " & objItem.IsLocked
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MonitorManufacturer: " & objItem.MonitorManufacturer
    Wscript.Echo "MonitorType: " & objItem.MonitorType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PixelsPerXLogicalInch: " & objItem.PixelsPerXLogicalInch
    Wscript.Echo "PixelsPerYLogicalInch: " & objItem.PixelsPerYLogicalInch
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ScreenHeight: " & objItem.ScreenHeight
    Wscript.Echo "ScreenWidth: " & objItem.ScreenWidth
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

