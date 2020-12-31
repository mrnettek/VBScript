' Description: Deletes all the network adapters for a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")
Set colNetworkAdapters = objVM.NetworkAdapters
For Each objAdapter in colNetworkAdapters
    objVM.RemoveNetworkAdapter(objAdapter)
Next

