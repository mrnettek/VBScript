On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_BinarySensor",,48)
For Each objItem in colItems
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
    Wscript.Echo "ExpectedReading: " & objItem.ExpectedReading
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InterpretationOfFalse: " & objItem.InterpretationOfFalse
    Wscript.Echo "InterpretationOfTrue: " & objItem.InterpretationOfTrue
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

