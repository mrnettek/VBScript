' Description: Adds a hard disk connection to a Virtual Machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")
Set objDrive = objVM.AddHardDiskConnection _
    ("c:\Virtual Machines\Windows 2000 Server Hard Disk.vhd",0,0,0)

