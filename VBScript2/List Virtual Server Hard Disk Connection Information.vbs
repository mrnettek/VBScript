' Description: Lists hard disk connection information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colHardDiskConnections = objVM.HardDiskConnections
    For Each objDrive in colHardDiskConnections
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Bus number: " & objDrive.BusNumber
        Wscript.Echo "Bus type: " & objDrive.BusType
        Wscript.Echo "Device number: " & objDrive.DeviceNumber
        Set objHardDisk = objDrive.HardDisk
        Wscript.Echo "Hard disk file: " & objHardDisk.File
        Wscript.Echo "Host drive identifier: " & _
            objHardDisk.HostDriveIdentifier
        Wscript.Echo "Host free disk space: " & objHardDisk.HostFreeDiskSpace
        Wscript.Echo "Host volume identifier: " & _
            objHardDisk.HostVolumeIdentifier
        Wscript.Echo "Size in guest: " & objHardDisk.SizeInGuest
        Wscript.Echo "Size on host: " & objHardDisk.SizeOnHost
        Wscript.Echo "Type: " & objHardDisk.Type
        Set objUndoDrive = objDrive.UndoHardDisk
        Wscript.Echo "Hard disk file: " & objUndoDrive.File
    Next  
Next

