' Description: Starts the DNS server service on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * From MicrosoftDNS_Server")

For Each objItem in colItems
    objItem.StartService()
Next

