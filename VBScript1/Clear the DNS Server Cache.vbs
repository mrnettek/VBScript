' Description: Clears the DNS server cache of resource records.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * From MicrosoftDNS_Cache")

For Each objItem in colItems
    objItem.ClearCache()
Next

