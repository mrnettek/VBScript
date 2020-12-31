' Description: Returns a list of all the Internet Connection Firewall port mappings on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_ConnectionPortMapping2")

For Each objItem in colItems
    Wscript.Echo "Connection: " & objItem.Connection
    Wscript.Echo "Enabled: " & objItem.Enabled
    Wscript.Echo "Name Active: " & objItem.NameActive
    Wscript.Echo "Protocol: " & objItem.Protocol
    Wscript.Echo "Target IP Address: " & objItem.TargetIPAddress
    Wscript.Echo "Target Name: " & objItem.TargetName
    Wscript.Echo "Target Port: " & objItem.TargetPort
Next

