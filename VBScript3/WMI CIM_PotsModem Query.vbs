On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_PotsModem",,48)
For Each objItem in colItems
    Wscript.Echo "AnswerMode: " & objItem.AnswerMode
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CompressionInfo: " & objItem.CompressionInfo
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CountriesSupported: " & objItem.CountriesSupported
    Wscript.Echo "CountrySelected: " & objItem.CountrySelected
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentPasswords: " & objItem.CurrentPasswords
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DialType: " & objItem.DialType
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorControlInfo: " & objItem.ErrorControlInfo
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "InactivityTimeout: " & objItem.InactivityTimeout
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MaxBaudRateToPhone: " & objItem.MaxBaudRateToPhone
    Wscript.Echo "MaxBaudRateToSerialPort: " & objItem.MaxBaudRateToSerialPort
    Wscript.Echo "MaxNumberOfPasswords: " & objItem.MaxNumberOfPasswords
    Wscript.Echo "ModulationScheme: " & objItem.ModulationScheme
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "RingsBeforeAnswer: " & objItem.RingsBeforeAnswer
    Wscript.Echo "SpeakerVolumeInfo: " & objItem.SpeakerVolumeInfo
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SupportsCallback: " & objItem.SupportsCallback
    Wscript.Echo "SupportsSynchronousConnect: " & objItem.SupportsSynchronousConnect
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
Next

