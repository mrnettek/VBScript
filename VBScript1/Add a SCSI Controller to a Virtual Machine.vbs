' Description: Adds a SCSI controller to a Virtual Machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")
objVM.AddSCSIController()

