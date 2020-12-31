On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Binary: " & objItem.Binary
    Wscript.Echo "Capabilities: " & objItem.Capabilities
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxBaudRate: " & objItem.MaxBaudRate
    Wscript.Echo "MaximumInputBufferSize: " & objItem.MaximumInputBufferSize
    Wscript.Echo "MaximumOutputBufferSize: " & objItem.MaximumOutputBufferSize
    Wscript.Echo "MaxNumberControlled: " & objItem.MaxNumberControlled
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OSAutoDiscovered: " & objItem.OSAutoDiscovered
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "ProtocolSupported: " & objItem.ProtocolSupported
    Wscript.Echo "ProviderType: " & objItem.ProviderType
    Wscript.Echo "SettableBaudRate: " & objItem.SettableBaudRate
    Wscript.Echo "SettableDataBits: " & objItem.SettableDataBits
    Wscript.Echo "SettableFlowControl: " & objItem.SettableFlowControl
    Wscript.Echo "SettableParity: " & objItem.SettableParity
    Wscript.Echo "SettableParityCheck: " & objItem.SettableParityCheck
    Wscript.Echo "SettableRLSD: " & objItem.SettableRLSD
    Wscript.Echo "SettableStopBits: " & objItem.SettableStopBits
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "Supports16BitMode: " & objItem.Supports16BitMode
    Wscript.Echo "SupportsDTRDSR: " & objItem.SupportsDTRDSR
    Wscript.Echo "SupportsElapsedTimeouts: " & objItem.SupportsElapsedTimeouts
    Wscript.Echo "SupportsIntTimeouts: " & objItem.SupportsIntTimeouts
    Wscript.Echo "SupportsParityCheck: " & objItem.SupportsParityCheck
    Wscript.Echo "SupportsRLSD: " & objItem.SupportsRLSD
    Wscript.Echo "SupportsRTSCTS: " & objItem.SupportsRTSCTS
    Wscript.Echo "SupportsSpecialCharacters: " & objItem.SupportsSpecialCharacters
    Wscript.Echo "SupportsXOnXOff: " & objItem.SupportsXOnXOff
    Wscript.Echo "SupportsXOnXOffSet: " & objItem.SupportsXOnXOffSet
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
Next

