On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Capabilities: " & objItem.Capabilities
    Wscript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CompressionMethod: " & objItem.CompressionMethod
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "DefaultBlockSize: " & objItem.DefaultBlockSize
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "Drive: " & objItem.Drive
    Wscript.Echo "DriveIntegrity: " & objItem.DriveIntegrity
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
    Wscript.Echo "FileSystemFlags: " & objItem.FileSystemFlags
    Wscript.Echo "FileSystemFlagsEx: " & objItem.FileSystemFlagsEx
    Wscript.Echo "Id: " & objItem.Id
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxBlockSize: " & objItem.MaxBlockSize
    Wscript.Echo "MaximumComponentLength: " & objItem.MaximumComponentLength
    Wscript.Echo "MaxMediaSize: " & objItem.MaxMediaSize
    Wscript.Echo "MediaLoaded: " & objItem.MediaLoaded
    Wscript.Echo "MediaType: " & objItem.MediaType
    Wscript.Echo "MfrAssignedRevisionLevel: " & objItem.MfrAssignedRevisionLevel
    Wscript.Echo "MinBlockSize: " & objItem.MinBlockSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NeedsCleaning: " & objItem.NeedsCleaning
    Wscript.Echo "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "RevisionLevel: " & objItem.RevisionLevel
    Wscript.Echo "SCSIBus: " & objItem.SCSIBus
    Wscript.Echo "SCSILogicalUnit: " & objItem.SCSILogicalUnit
    Wscript.Echo "SCSIPort: " & objItem.SCSIPort
    Wscript.Echo "SCSITargetId: " & objItem.SCSITargetId
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TransferRate: " & objItem.TransferRate
    Wscript.Echo "VolumeName: " & objItem.VolumeName
    Wscript.Echo "VolumeSerialNumber: " & objItem.VolumeSerialNumber
Next

