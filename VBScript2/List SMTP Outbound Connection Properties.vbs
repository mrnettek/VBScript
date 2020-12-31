' Description: Lists the default SMTP outbound connection metabase properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Maximum Out Connections: " & _
        objItem.MaxOutConnections
    Wscript.Echo "Maximum Out Connections Per Domain: " & _
        objItem.MaxOutConnectionsPerDomain
    Wscript.Echo "Remote SMTP Port: " & objItem.RemoteSmtpPort
    Wscript.Echo "Remote Timeout: " & objItem.RemoteTimeout
Next

