' Description: Returns information about the ISDN resource records on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_ISDNType")

For Each objItem in colItems
    Wscript.Echo "Owner Name: " & objItem.OwnerName
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "DNS Serve Name: " & objItem.DnsServerName
    Wscript.Echo "Domain Name: " & objItem.DomainName
    Wscript.Echo "ISDN Number: " & objItem.ISDNNumber
    Wscript.Echo "Record Class: " & objItem.RecordClass
    Wscript.Echo "Record Data: " & objItem.RecordData
    Wscript.Echo "SubAddress: " & objItem.SubAddress
    Wscript.Echo "Text Representation: " & objItem.TextRepresentation
    Wscript.Echo "Timestamp: " & objItem.Timestamp
    Wscript.Echo "Time-to-Live: " & objItem.TTL
    Wscript.Echo
Next

