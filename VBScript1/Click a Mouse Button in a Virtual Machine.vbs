' Description: Clicks a mouse button inside a virtual machine named Windows 2000.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objMouse = objVM.Mouse
errReturn = objMouse.Click(2)

