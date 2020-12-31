' Description: Demonstration script that modifies global SMTP metabase property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    objItem.ConnectionTimeout = 1200
    objItem.DontLog = True
    objItem.MaxConnections = 10000
    objItem.Put_
Next

