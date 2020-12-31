' Description: Detaches all the network adapters found on a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set colNetworkAdapters = objVM.NetworkAdapters
For Each objNetworkAdapter in colNetworkAdapters
    errReturn = objNetworkAdapter.DetachFromVirtualNetwork()
Next

