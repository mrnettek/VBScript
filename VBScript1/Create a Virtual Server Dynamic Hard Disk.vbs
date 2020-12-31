' Description: Creates a new Virtual Server dynamic hard disk.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.CreateDynamicVirtualHardDisk _
    ("C:\Virtual Machines\Disks\Scripted_HardDisk.vhd", 20)

