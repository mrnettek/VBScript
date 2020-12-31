' Description: Lists mouse information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set objMouse = objVM.Mouse
    Wscript.Echo objVM.Name
    Wscript.Echo "Horizontal position: " & objMouse.HorizontalPosition
    Wscript.Echo "Scroll wheel position: " & objMouse.ScrollWheelPosition
    Wscript.Echo "Using absolute coordinates position: " & _
        objMouse.UsingAbsoluteCoordinatesPosition
    Wscript.Echo "Vertical position: " & objMouse.VerticalPosition
    Wscript.Echo
Next

