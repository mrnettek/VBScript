On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Printer",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "AvailableJobSheets: " & objItem.AvailableJobSheets
    Wscript.Echo "Capabilities: " & objItem.Capabilities
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CharSetsSupported: " & objItem.CharSetsSupported
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentCapabilities: " & objItem.CurrentCapabilities
    Wscript.Echo "CurrentCharSet: " & objItem.CurrentCharSet
    Wscript.Echo "CurrentLanguage: " & objItem.CurrentLanguage
    Wscript.Echo "CurrentMimeType: " & objItem.CurrentMimeType
    Wscript.Echo "CurrentNaturalLanguage: " & objItem.CurrentNaturalLanguage
    Wscript.Echo "CurrentPaperType: " & objItem.CurrentPaperType
    Wscript.Echo "DefaultCapabilities: " & objItem.DefaultCapabilities
    Wscript.Echo "DefaultCopies: " & objItem.DefaultCopies
    Wscript.Echo "DefaultLanguage: " & objItem.DefaultLanguage
    Wscript.Echo "DefaultMimeType: " & objItem.DefaultMimeType
    Wscript.Echo "DefaultNumberUp: " & objItem.DefaultNumberUp
    Wscript.Echo "DefaultPaperType: " & objItem.DefaultPaperType
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DetectedErrorState: " & objItem.DetectedErrorState
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorInformation: " & objItem.ErrorInformation
    Wscript.Echo "HorizontalResolution: " & objItem.HorizontalResolution
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "JobCountSinceLastReset: " & objItem.JobCountSinceLastReset
    Wscript.Echo "LanguagesSupported: " & objItem.LanguagesSupported
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "MarkingTechnology: " & objItem.MarkingTechnology
    Wscript.Echo "MaxCopies: " & objItem.MaxCopies
    Wscript.Echo "MaxNumberUp: " & objItem.MaxNumberUp
    Wscript.Echo "MaxSizeSupported: " & objItem.MaxSizeSupported
    Wscript.Echo "MimeTypesSupported: " & objItem.MimeTypesSupported
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NaturalLanguagesSupported: " & objItem.NaturalLanguagesSupported
    Wscript.Echo "PaperSizesSupported: " & objItem.PaperSizesSupported
    Wscript.Echo "PaperTypesAvailable: " & objItem.PaperTypesAvailable
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "PrinterStatus: " & objItem.PrinterStatus
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
    Wscript.Echo "VerticalResolution: " & objItem.VerticalResolution
Next

