On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_MemoryArray",,48)
For Each objItem in colItems
    Wscript.Echo "Access: " & objItem.Access
    Wscript.Echo "AdditionalErrorData: " & objItem.AdditionalErrorData
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "BlockSize: " & objItem.BlockSize
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CorrectableError: " & objItem.CorrectableError
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "EndingAddress: " & objItem.EndingAddress
    Wscript.Echo "ErrorAccess: " & objItem.ErrorAccess
    Wscript.Echo "ErrorAddress: " & objItem.ErrorAddress
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorData: " & objItem.ErrorData
    Wscript.Echo "ErrorDataOrder: " & objItem.ErrorDataOrder
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorGranularity: " & objItem.ErrorGranularity
    Wscript.Echo "ErrorInfo: " & objItem.ErrorInfo
    Wscript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
    Wscript.Echo "ErrorResolution: " & objItem.ErrorResolution
    Wscript.Echo "ErrorTime: " & objItem.ErrorTime
    Wscript.Echo "ErrorTransferSize: " & objItem.ErrorTransferSize
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
    Wscript.Echo "OtherErrorDescription: " & objItem.OtherErrorDescription
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Purpose: " & objItem.Purpose
    Wscript.Echo "StartingAddress: " & objItem.StartingAddress
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemLevelAddress: " & objItem.SystemLevelAddress
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

