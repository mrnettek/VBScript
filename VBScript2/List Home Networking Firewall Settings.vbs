' Description: Enumerates the inbound and outbound configuration settings for Internet Connection Firewall.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery("Select * from HNet_FwIcmpSettings")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Allow Inbound Echo request: " & _
        objItem.AllowInboundEchoRequest
    Wscript.Echo "Allow Inbound Mask Requestt: " & _
        objItem.AllowInboundMaskRequest
    Wscript.Echo "Allow Inbound Router Request: " & _
        objItem.AllowInboundRouterRequest
    Wscript.Echo "Allow Inbound Timestamp Request: " & _
        objItem.AllowInboundTimestampRequest
    Wscript.Echo "Allow Outbound Destination Unreachable: " & _
        objItem.AllowOutboundDestinationUnreachable
    Wscript.Echo "Allow Outbound Parameter Problem: " & _
        objItem.AllowOutboundParameterProblem
    Wscript.Echo "Allow Outbound Source Quench: " & _
        objItem.AllowOutboundSourceQuench
    Wscript.Echo "Allow Outbound Time Exceeded: " & _
        objItem.AllowOutboundTimeExceeded
    Wscript.Echo "Allow redirect: " & objItem.AllowRedirect
Next

