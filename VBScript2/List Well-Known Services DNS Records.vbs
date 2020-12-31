' Description: Returns information about all the well-known services (WKS) records on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_WKSType")

For Each objItem in colItems
    Wscript.Echo "Owner Name: " & objItem.OwnerName
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "DNS Server Name: " & objItem.DnsServerName
    Wscript.Echo "Domain Name: " & objItem.DomainName
    Wscript.Echo "Internet Address: " & objItem.InternetAddress
    Wscript.Echo "IP Protocol: " & objItem.IPProtocol
    Wscript.Echo "Record Class: " & objItem.RecordClass
    Wscript.Echo "Record Data: " & objItem.RecordData
    Wscript.Echo "Services: " & objItem.Services
    Wscript.Echo "Text Representation: " & objItem.TextRepresentation
    Wscript.Echo "Timestamp: " & objItem.Timestamp
    Wscript.Echo "Time-to-Live: " & objItem.TTL
    Wscript.Echo
Next

