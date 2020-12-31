' Description: Retrieves the properties for all the physical disk drives installed on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colDiskDrives = objWMIService.ExecQuery _    
    ("Select * from Win32_DiskDrive")

For each objDiskDrive in colDiskDrives
    Wscript.Echo "Bytes Per Sector: " & vbTab &  _
        objDiskDrive.BytesPerSector        
    For i = Lbound(objDiskDrive.Capabilities) to _
        Ubound(objDiskDrive.Capabilities)
        Wscript.Echo "Capabilities: " & vbTab &  _
            objDiskDrive.Capabilities(i)
    Next    
    Wscript.Echo "Caption: " & vbTab &  objDiskDrive.Caption
    Wscript.Echo "Device ID: " & vbTab &  objDiskDrive.DeviceID
    Wscript.Echo "Index: " & vbTab &  objDiskDrive.Index
    Wscript.Echo "Interface Type: " & vbTab & objDiskDrive.InterfaceType
    Wscript.Echo "Manufacturer: " & vbTab & objDiskDrive.Manufacturer
    Wscript.Echo "Media Loaded: " & vbTab  & objDiskDrive.MediaLoaded
    Wscript.Echo "Media Type: " & vbTab &  objDiskDrive.MediaType
    Wscript.Echo "Model: " & vbTab &  objDiskDrive.Model
    Wscript.Echo "Name: " & vbTab &  objDiskDrive.Name
    Wscript.Echo "Partitions: " & vbTab & objDiskDrive.Partitions
    Wscript.Echo "PNP DeviceID: " & vbTab &  objDiskDrive.PNPDeviceID
    Wscript.Echo "SCSI Bus: " & vbTab &  objDiskDrive.SCSIBus
    Wscript.Echo "SCSI Logical Unit: " & vbTab &  _
        objDiskDrive.SCSILogicalUnit
    Wscript.Echo "SCSI Port: " & vbTab &  objDiskDrive.SCSIPort
    Wscript.Echo "SCSI TargetId: " & vbTab &  objDiskDrive.SCSITargetId    
    Wscript.Echo "Sectors Per Track: " & vbTab &  _
        objDiskDrive.SectorsPerTrack        
    Wscript.Echo "Signature: " & vbTab &  objDiskDrive.Signature          
    Wscript.Echo "Size: " & vbTab &  objDiskDrive.Size     
    Wscript.Echo "Status: " & vbTab &  objDiskDrive.Status         
    Wscript.Echo "Total Cylinders: " & vbTab &  _
        objDiskDrive.TotalCylinders         
    Wscript.Echo "Total Heads: " & vbTab &  objDiskDrive.TotalHeads    
    Wscript.Echo "Total Sectors: " & vbTab &  objDiskDrive.TotalSectors
    Wscript.Echo "Total Tracks: " & vbTab &  objDiskDrive.TotalTracks
    Wscript.Echo "Tracks Per Cylinder: " & vbTab &  _
        objDiskDrive.TracksPerCylinder  
Next

