' Description: Returns information about the IP route tables configured on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_IP4RouteTable")

For Each objItem in colItems
    Wscript.Echo "Age: " & objItem.Age
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Destination: " & objItem.Destination
    Wscript.Echo "Information: " & objItem.Information
    Wscript.Echo "Interface Index: " & objItem.InterfaceIndex
    Wscript.Echo "Mask: " & objItem.Mask
    Wscript.Echo "Metric 1: " & objItem.Metric1
    Wscript.Echo "Metric 2: " & objItem.Metric2
    Wscript.Echo "Metric 3: " & objItem.Metric3
    Wscript.Echo "Metric 4: " & objItem.Metric4
    Wscript.Echo "Metric 5: " & objItem.Metric5
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Next Hop: " & objItem.NextHop
    Wscript.Echo "Protocol: " & objItem.Protocol
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo
Next

