' Description: Demonstration script that modifies global SMTP outbound connection metabase properties on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    objItem.MaxOutConnections = 500
    objItem.MaxOutConnectionsPerDomain = 250
    objItem.RemoteSmtpPort = 25
    objItem.RemoteTimeout = 900
    objItem.Put_
Next

