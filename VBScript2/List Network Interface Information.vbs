' Description: Returns information about the network interfaces on a cluster server.


On Error Resume Next 

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")

Set colItems = objWMIService.ExecQuery _
    ("Select * from MSCluster_NetworkInterface")

For Each objItem in colItems
    Wscript.Echo "Adapter: " & objItem.Adapter
    Wscript.Echo "Address: " & objItem.Address
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Characteristics: " & objItem.Characteristics
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "Identifying descriptions: " & objItem.IdentifyingDescriptions
    Wscript.Echo "Installation date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Network: " & objItem.Network
    Wscript.Echo "Other identifying info: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "Power-on hours: " & objItem.PowerOnHours
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Status info: " & objItem.StatusInfo
    Wscript.Echo "System name: " & objItem.SystemName
    Wscript.Echo "Total power-on hours: " & objItem.TotalPowerOnHours
    Wscript.Echo
Next

