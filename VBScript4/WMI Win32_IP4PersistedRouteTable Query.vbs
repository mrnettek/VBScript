On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_IP4PersistedRouteTable",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Destination: " & objItem.Destination
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Mask: " & objItem.Mask
    Wscript.Echo "Metric1: " & objItem.Metric1
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NextHop: " & objItem.NextHop
    Wscript.Echo "Status: " & objItem.Status
Next

