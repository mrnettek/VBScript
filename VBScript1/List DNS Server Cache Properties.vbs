' Description: Returns information about the DNS cache on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_Cache")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "DNS Server Name: " & objItem.DNSServerName
    Wscript.Echo
Next

