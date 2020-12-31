' Description: Creates a new Virtual Server fixed hard disk.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.CreateFixedVirtualHardDisk _
    ("C:\Virtual Machines\Disks\Fixed_HardDisk.vhd", 20)

