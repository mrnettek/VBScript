' Description: Lists DVD information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colDVDDrives = objVM.DVDROMDrives
    For Each objDrive in colDVDDrives
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Attachment: " & objDrive.Attachment
        Wscript.Echo "Bus number: " & objDrive.BusNumber
        Wscript.Echo "Bus type: " & objDrive.BusType
        Wscript.Echo "Device number: " & objDrive.DeviceNumber
        Wscript.Echo "Host drive letter: " & objDrive.HostDriveLetter
        Wscript.Echo "Image file: " & objDrive.ImageFile
        Wscript.Echo
    Next
Next

