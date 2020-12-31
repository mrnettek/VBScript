On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_CacheMemory",,48)
For Each objItem in colItems
    Wscript.Echo "Access: " & objItem.Access
    Wscript.Echo "AdditionalErrorData: " & objItem.AdditionalErrorData
    Wscript.Echo "Associativity: " & objItem.Associativity
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "BlockSize: " & objItem.BlockSize
    Wscript.Echo "CacheSpeed: " & objItem.CacheSpeed
    Wscript.Echo "CacheType: " & objItem.CacheType
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CorrectableError: " & objItem.CorrectableError
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CurrentSRAM: " & objItem.CurrentSRAM
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "EndingAddress: " & objItem.EndingAddress
    Wscript.Echo "ErrorAccess: " & objItem.ErrorAccess
    Wscript.Echo "ErrorAddress: " & objItem.ErrorAddress
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorCorrectType: " & objItem.ErrorCorrectType
    Wscript.Echo "ErrorData: " & objItem.ErrorData
    Wscript.Echo "ErrorDataOrder: " & objItem.ErrorDataOrder
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorInfo: " & objItem.ErrorInfo
    Wscript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
    Wscript.Echo "ErrorResolution: " & objItem.ErrorResolution
    Wscript.Echo "ErrorTime: " & objItem.ErrorTime
    Wscript.Echo "ErrorTransferSize: " & objItem.ErrorTransferSize
    Wscript.Echo "FlushTimer: " & objItem.FlushTimer
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InstalledSize: " & objItem.InstalledSize
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Level: " & objItem.Level
    Wscript.Echo "LineSize: " & objItem.LineSize
    Wscript.Echo "Location: " & objItem.Location
    Wscript.Echo "MaxCacheSize: " & objItem.MaxCacheSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
    Wscript.Echo "OtherErrorDescription: " & objItem.OtherErrorDescription
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "Purpose: " & objItem.Purpose
    Wscript.Echo "ReadPolicy: " & objItem.ReadPolicy
    Wscript.Echo "ReplacementPolicy: " & objItem.ReplacementPolicy
    Wscript.Echo "StartingAddress: " & objItem.StartingAddress
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SupportedSRAM: " & objItem.SupportedSRAM
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemLevelAddress: " & objItem.SystemLevelAddress
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "WritePolicy: " & objItem.WritePolicy
Next

