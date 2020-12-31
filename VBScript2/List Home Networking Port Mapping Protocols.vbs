' Description: Enumerates all the Internet Connection Firewall port mapping protocols on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_PortMappingProtocol")

For Each objItem in colItems
    Wscript.Echo "Built In: " & objItem.BuiltIn
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "IP Protocol: " & objItem.IPProtocol
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Port: " & objItem.Port
Next

