' Description: Modifies the scheduling parameters for a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objAccountant = objVM.Accountant
errReturn = objAccountant.SetSchedulingParameters(99,99,99)

