' Description: Lists floppy drive information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colFloppyDrives = objVM.FloppyDrives
    For Each objDrive in colFloppyDrives
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Attachment: " & objDrive.Attachment
        Wscript.Echo "Drive number: " & objDrive.DriveNumber
        Wscript.Echo "Host drive letter: " & objDrive.HostDriveLetter
        Wscript.Echo "Image file: " & objDrive.ImageFile
        Wscript.Echo
    Next
Next

