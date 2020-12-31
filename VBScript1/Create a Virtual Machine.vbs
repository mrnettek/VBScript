' Description: Creates a virtual machine named Script Machine.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.CreateVirtualMachine("Scripted Machine", _
    "C:\Scripts\Shared Virtual Machines\Scripted")

