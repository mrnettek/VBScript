' Description: Deletes a virtual machine named Scripted Machine.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Scripted Machine")

errReturn = objVS.DeleteVirtualMachine(objVM)

