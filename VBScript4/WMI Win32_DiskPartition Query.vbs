On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskPartition",,48)
For Each objItem in colItems
    Wscript.Echo "Access: " & objItem.Access
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "BlockSize: " & objItem.BlockSize
    Wscript.Echo "Bootable: " & objItem.Bootable
    Wscript.Echo "BootPartition: " & objItem.BootPartition
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
    Wscript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DiskIndex: " & objItem.DiskIndex
    Wscript.Echo "ErrorCleared: " & objItem.ErrorCleared
    Wscript.Echo "ErrorDescription: " & objItem.ErrorDescription
    Wscript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
    Wscript.Echo "HiddenSectors: " & objItem.HiddenSectors
    Wscript.Echo "Index: " & objItem.Index
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastErrorCode: " & objItem.LastErrorCode
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
    Wscript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "PrimaryPartition: " & objItem.PrimaryPartition
    Wscript.Echo "Purpose: " & objItem.Purpose
    Wscript.Echo "RewritePartition: " & objItem.RewritePartition
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "StartingOffset: " & objItem.StartingOffset
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusInfo: " & objItem.StatusInfo
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "Type: " & objItem.Type
Next

