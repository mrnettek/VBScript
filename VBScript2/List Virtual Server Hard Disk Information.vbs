' Description: Lists the disk type for a Virtual Server hard disk.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objHardDisk = objVS.GetHardDisk _
    ("C:\Virtual Machines\Disks\Windows 2000 Server Hard Disk.vhd")
Wscript.Echo "Hard disk type: " & objHardDisk.Type

