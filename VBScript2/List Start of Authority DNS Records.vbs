' Description: Returns information about all the start of authority (SOA) records on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_SOAType")

For Each objItem in colItems
    Wscript.Echo "Owner Name: " & objItem.OwnerName
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "DNS Server Name: " & objItem.DnsServerName
    Wscript.Echo "Domain Name: " & objItem.DomainName
    Wscript.Echo "Expire Limit: " & objItem.ExpireLimit
    Wscript.Echo "Minimum Time-to-Live: " & objItem.MinimumTTL
    Wscript.Echo "Primary Server: " & objItem.PrimaryServer
    Wscript.Echo "Record Class: " & objItem.RecordClass
    Wscript.Echo "Record Data: " & objItem.RecordData
    Wscript.Echo "Refresh Interval: " & objItem.RefreshInterval
    Wscript.Echo "Responsible Party: " & objItem.ResponsibleParty
    Wscript.Echo "Retry Delay: " & objItem.RetryDelay
    Wscript.Echo "Serial Number: " & objItem.SerialNumber
    Wscript.Echo "Text Representation: " & objItem.TextRepresentation
    Wscript.Echo "Time-to-Live: " & objItem.TTL
    Wscript.Echo
Next

