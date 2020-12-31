' Description: Lists display information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set objDisplay = objVM.Display
    Wscript.Echo objVM.Name
    Wscript.Echo "Height: " & objDisplay.Height
    Wscript.Echo "Video mode: " & objDisplay.VideoMode
    Wscript.Echo "Width: " & objDisplay.Width
    Wscript.Echo
Next

