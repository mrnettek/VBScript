' Description: Lists serial port information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colPorts = objVM.SerialPorts
    For Each objPort in colPorts
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Connect immediately: " & objPort.ConnectImmediately
        Wscript.Echo "Name: " & objPort.Name
        Wscript.Echo "Type: " & objPort.Type
        Wscript.Echo
    Next
Next

