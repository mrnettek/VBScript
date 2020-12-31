' Description: Lists the properties for all the logical disk drives on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")

For each objDisk in colDisks
    Wscript.Echo "Compressed: " & objDisk.Compressed  
    Wscript.Echo "Description: " & objDisk.Description       
    Wscript.Echo "DeviceID: " & objDisk.DeviceID      
    Wscript.Echo "DriveType: " & objDisk.DriveType    
    Wscript.Echo "FileSystem: " & objDisk.FileSystem  
    Wscript.Echo "FreeSpace: " & objDisk.FreeSpace    
    Wscript.Echo "MediaType: " & objDisk.MediaType    
    Wscript.Echo "Name: " & objDisk.Name      
    Wscript.Echo "QuotasDisabled: " & objDisk.QuotasDisabled
    Wscript.Echo "QuotasIncomplete: " & objDisk.QuotasIncomplete
    Wscript.Echo "QuotasRebuilding: " & objDisk.QuotasRebuilding
    Wscript.Echo "Size: " & objDisk.Size      
    Wscript.Echo "SupportsDiskQuotas: " & _
        objDisk.SupportsDiskQuotas      
    Wscript.Echo "SupportsFileBasedCompression: " & _
        objDisk.SupportsFileBasedCompression   
    Wscript.Echo "SystemName: " & objDisk.SystemName  
    Wscript.Echo "VolumeDirty: " & objDisk.VolumeDirty       
    Wscript.Echo "VolumeName: " & objDisk.VolumeName  
    Wscript.Echo "VolumeSerialNumber: " & _
        objDisk.VolumeSerialNumber      
Next

