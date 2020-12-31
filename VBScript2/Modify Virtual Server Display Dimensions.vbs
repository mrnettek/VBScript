' Description: Modifies the display dimensions for a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objDisplay = objVM.Display
errReturn = objDisplay.SetDimensions(800,600)

