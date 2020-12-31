' Description: Stops the DNS server service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * From MicrosoftDNS_Server")

For Each objItem in colItems
    objItem.StopService()
Next

