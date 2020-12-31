On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Tachometer",,48)
For Each objItem in colItems
    Wscript.Echo "Accuracy: " & objItem.Accuracy
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentReading: " & objItem.CurrentReading
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IsLinear: " & objItem.IsLinear
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "LowerThresholdCritical: " & objItem.LowerThresholdCritical
    Wscript.Echo "LowerThresholdFatal: " & objItem.LowerThresholdFatal
    Wscript.Echo "LowerThresholdNonCritical: " & objItem.LowerThresholdNonCritical
    Wscript.Echo "MaxReadable: " & objItem.MaxReadable
    Wscript.Echo "MinReadable: " & objItem.MinReadable
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NominalReading: " & objItem.NominalReading
    Wscript.Echo "NormalMax: " & objItem.NormalMax
    Wscript.Echo "NormalMin: " & objItem.NormalMin
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Resolution: " & objItem.Resolution
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "Tolerance: " & objItem.Tolerance
    Wscript.Echo "UpperThresholdCritical: " & objItem.UpperThresholdCritical
    Wscript.Echo "UpperThresholdFatal: " & objItem.UpperThresholdFatal
    Wscript.Echo "UpperThresholdNonCritical: " & objItem.UpperThresholdNonCritical
Next

