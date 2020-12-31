On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "BatteryRechargeTime: " & objItem.BatteryRechargeTime
    Wscript.Echo "BatteryStatus: " & objItem.BatteryStatus
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Chemistry: " & objItem.Chemistry
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DesignCapacity: " & objItem.DesignCapacity
    Wscript.Echo "DesignVoltage: " & objItem.DesignVoltage
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "EstimatedChargeRemaining: " & objItem.EstimatedChargeRemaining
    Wscript.Echo "EstimatedRunTime: " & objItem.EstimatedRunTime
    Wscript.Echo "ExpectedBatteryLife: " & objItem.ExpectedBatteryLife
    Wscript.Echo "ExpectedLife: " & objItem.ExpectedLife
    Wscript.Echo "FullChargeCapacity: " & objItem.FullChargeCapacity
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxRechargeTime: " & objItem.MaxRechargeTime
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "SmartBatteryVersion: " & objItem.SmartBatteryVersion
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOnBattery: " & objItem.TimeOnBattery
    Wscript.Echo "TimeToFullCharge: " & objItem.TimeToFullCharge
Next

