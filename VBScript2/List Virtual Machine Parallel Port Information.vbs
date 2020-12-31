' Description: Lists parallel port information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colPorts = objVM.ParallelPorts
    For Each objPort in colPorts
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Name: " & objPort.Name
        Wscript.Echo
    Next
Next

