' Description: Lists SCSI controller information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colSCSIControllers = objVM.SCSIControllers
    For Each objController in colSCSIControllers
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "ID: " & objController.ID
        Wscript.Echo "Is bus shared: " & objController.IsBusShared
        Wscript.Echo "SCSI ID: " & objController.SCSIID
        Wscript.Echo
    Next
Next

