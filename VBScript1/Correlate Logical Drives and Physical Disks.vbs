strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colDiskDrives = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
 
For Each objDrive In colDiskDrives
    Wscript.Echo "Physical Disk: " & objDrive.Caption & " -- " & objDrive.DeviceID 
    strDeviceID = Replace(objDrive.DeviceID, "\", "\\")
    Set colPartitions = objWMIService.ExecQuery _
        ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & _
            strDeviceID & """} WHERE AssocClass = " & _
                "Win32_DiskDriveToDiskPartition")
 
    For Each objPartition In colPartitions
        Wscript.Echo "Disk Partition: " & objPartition.DeviceID
        Set colLogicalDisks = objWMIService.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & _
                objPartition.DeviceID & """} WHERE AssocClass = " & _
                    "Win32_LogicalDiskToPartition")
 
        For Each objLogicalDisk In colLogicalDisks
            Wscript.Echo "Logical Disk: " & objLogicalDisk.DeviceID
        Next
        Wscript.Echo
    Next
    Wscript.Echo
Next
  


