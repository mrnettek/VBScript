' Description: Retrieves information about mapped network drives. The information returned is similar to that available through the Win32_LogicalDisk class, which retrieves information about the logical disks found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")

For Each objItem in colItems
    Wscript.Echo "Compressed: " & objItem.Compressed
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "File System: " & objItem.FileSystem
    Wscript.Echo "Free Space: " & objItem.FreeSpace
    Wscript.Echo "Maximum Component Length: " & objItem.MaximumComponentLength
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Provider Name: " & objItem.ProviderName
    Wscript.Echo "Session ID: " & objItem.SessionID
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "Supports Disk Quotas: " & objItem.SupportsDiskQuotas
    Wscript.Echo "Supports File-Based Compression: " & _
        objItem.SupportsFileBasedCompression
    Wscript.Echo "Volume Name: " & objItem.VolumeName
    Wscript.Echo "Volume Serial Number: " & objItem.VolumeSerialNumber
    Wscript.Echo
Next

