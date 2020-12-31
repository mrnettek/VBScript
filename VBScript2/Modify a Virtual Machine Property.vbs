' Description: Modifies the Notes property for a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

objVM.Notes = "This is a Windows 2000 virtual machine."

