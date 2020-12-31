' Description: Returns information about all the signature (SIG) resource records on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_SIGType")

For Each objItem in colItems
    Wscript.Echo "Owner Name: " & objItem.OwnerName
    Wscript.Echo "Algorithm: " & objItem.Algorithm
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "DNS Server Name: " & objItem.DnsServerName
    Wscript.Echo "Domain Name: " & objItem.DomainName
    Wscript.Echo "Key Tag: " & objItem.KeyTag
    Wscript.Echo "Labels: " & objItem.Labels
    Wscript.Echo "Original Time-to-Live: " & objItem.OriginalTTL
    Wscript.Echo "Record Class: " & objItem.RecordClass
    Wscript.Echo "Record Data: " & objItem.RecordData
    Wscript.Echo "Signature: " & objItem.Signature
    Wscript.Echo "Signature Expiration: " & objItem.SignatureExpiration
    Wscript.Echo "Signature Inception: " & objItem.SignatureInception
    Wscript.Echo "Signer Name: " & objItem.SignerName
    Wscript.Echo "Text Representation: " & objItem.TextRepresentation
    Wscript.Echo "Timestamp: " & objItem.Timestamp
    Wscript.Echo "Time-to-Live: " & objItem.TTL
    Wscript.Echo "Type Covered: " & objItem.TypeCovered
    Wscript.Echo
Next

