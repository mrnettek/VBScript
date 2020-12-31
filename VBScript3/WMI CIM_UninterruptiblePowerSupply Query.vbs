On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_UninterruptiblePowerSupply",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveInputVoltage: " & objItem.ActiveInputVoltage
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "EstimatedChargeRemaining: " & objItem.EstimatedChargeRemaining
    Wscript.Echo "EstimatedRunTime: " & objItem.EstimatedRunTime
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IsSwitchingSupply: " & objItem.IsSwitchingSupply
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Range1InputFrequencyHigh: " & objItem.Range1InputFrequencyHigh
    Wscript.Echo "Range1InputFrequencyLow: " & objItem.Range1InputFrequencyLow
    Wscript.Echo "Range1InputVoltageHigh: " & objItem.Range1InputVoltageHigh
    Wscript.Echo "Range1InputVoltageLow: " & objItem.Range1InputVoltageLow
    Wscript.Echo "Range2InputFrequencyHigh: " & objItem.Range2InputFrequencyHigh
    Wscript.Echo "Range2InputFrequencyLow: " & objItem.Range2InputFrequencyLow
    Wscript.Echo "Range2InputVoltageHigh: " & objItem.Range2InputVoltageHigh
    Wscript.Echo "Range2InputVoltageLow: " & objItem.Range2InputVoltageLow
    Wscript.Echo "RemainingCapacityStatus: " & objItem.RemainingCapacityStatus
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOnBackup: " & objItem.TimeOnBackup
    Wscript.Echo "TotalOutputPower: " & objItem.TotalOutputPower
    Wscript.Echo "TypeOfRangeSwitching: " & objItem.TypeOfRangeSwitching
Next

