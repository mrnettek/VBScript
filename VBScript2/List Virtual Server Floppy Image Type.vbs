' Description: Lists the image type for a virtual server floppy image.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
strImageType = objVS.GetFloppyDiskImageType _
    ("C:\Virtual Machines\Images\Dos Virtual Machine Additions.vfd")
Wscript.Echo "Image type: " & strImageType

