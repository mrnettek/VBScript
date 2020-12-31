' Description: Locates the security entry for a Virtual Server user named bob on a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objSecurity = objVM.Security
Set objEntry = objSecurity.FindEntry("bob",0)

Wscript.Echo objEntry.Name
Wscript.Echo objEntry.SID

