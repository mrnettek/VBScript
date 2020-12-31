' Description: Returns a list of all the Plug and Play devices installed in a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")

For Each objItem in colItems
    Wscript.Echo "Class GUID: " & objItem.ClassGuid
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Service: " & objItem.Service
Next

