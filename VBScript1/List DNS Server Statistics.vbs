' Description: Returns statistics collected on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_Statistic")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Collection Name: " & objItem.CollectionName
    Wscript.Echo "Collection ID: " & objItem.CollectionId
    Wscript.Echo "DNS Server Name: " & objItem.DnsServerName
    Wscript.Echo "String Value: " & objItem.StringValue
    Wscript.Echo "Value: " & objItem.Value
    Wscript.Echo
Next

