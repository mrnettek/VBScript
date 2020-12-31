' Description: Returns mouse button status for a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objMouse = objVM.Mouse
errReturn = objMouse.GetButton(1)
Wscript.Echo errReturn

