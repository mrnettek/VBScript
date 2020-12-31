' Description: Lists keyboard information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set objKeyboard = objVM.Keyboard
    Wscript.Echo objVM.Name
    Wscript.Echo "Has exclusive access: " & objKeyboard.HasExclusiveAccess
    Wscript.Echo
Next

