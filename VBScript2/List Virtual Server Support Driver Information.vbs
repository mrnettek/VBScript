' Description: Lists information about support drives for Virtual Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colDrivers = objVS.SupportDrivers

For Each objDriver in colDrivers
    Wscript.Echo "Date: " & objDriver.Date
    Wscript.Echo "Description: " & objDriver.Description
    Wscript.Echo "Manufacturer: " & objDriver.Manufacturer
    Wscript.Echo "Provider: " & objDriver.Provider
    Wscript.Echo "Type: " & objDriver.Type
    Wscript.Echo "Version: " & objDriver.Version
    Wscript.Echo
Next

