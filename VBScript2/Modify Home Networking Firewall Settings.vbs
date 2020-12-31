' Description: Disables the ability of the Internet Connection Firewall to allow inbound echo requests. To enable the ability to allow inbound echo requests, simply set the value of the property AlllowInboundEchoRequest to True.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery("Select * from HNet_FwIcmpSettings")

For Each objItem in colItems
    objItem.AllowInboundEchoRequest = False
    objItem.Put_
Next

