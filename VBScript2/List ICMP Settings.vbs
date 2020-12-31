' Description: Lists ICMP settings for the Windows Firewall current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objICMPSettings = objPolicy.ICMPSettings

Wscript.Echo "Allow inbound echo request: " & _
    objICMPSettings.AllowInboundEchoRequest
Wscript.Echo "Allow inbound mask request: " & _
    objICMPSettings.AllowInboundMaskRequest
Wscript.Echo "Allow inbound router request: " & _
    objICMPSettings.AllowInboundRouterRequest
Wscript.Echo "Allow inbound timestamp request: " & _
    objICMPSettings.AllowInboundTimestampRequest
Wscript.Echo "Allow outbound destination unreachable: " & _
    objICMPSettings.AllowOutboundDestinationUnreachable
Wscript.Echo "Allow outbound packet too big: " & _
    objICMPSettings.AllowOutboundPacketTooBig
Wscript.Echo "Allow outbound parameter problem: " & _
    objICMPSettings.AllowOutboundParameterProblem
Wscript.Echo "Allow outbound source quench: " & _
    objICMPSettings.AllowOutboundSourceQuench
Wscript.Echo "Allow outbound time exceeded: " & _
    objICMPSettings.AllowOutboundTimeExceeded
Wscript.Echo "Allow redirect: " & objICMPSettings.AllowRedirect

