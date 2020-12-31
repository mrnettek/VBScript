' Description: Lists network adapters for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colNetworkAdapters = objVM.NetworkAdapters
    For Each objAdapter in colNetworkAdapters
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Network adapter ID: " & objAdapter.ID
        Wscript.Echo "Ethernet address: " & objAdapter.EthernetAddress
        Wscript.Echo "Is ethernet address dynamic: " & _
            objAdapter.IsEthernetAddressDynamic
        Wscript.Echo "Virtual machine: " & objAdapter.VirtualMachine
        Wscript.Echo "Virtual network: " & objAdapter.VirtualNetwork
        Wscript.Echo
    Next
Next

