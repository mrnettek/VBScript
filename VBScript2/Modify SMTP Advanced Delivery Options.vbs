' Description: Demonstration script that modifies global SMTP advanced delivery metabase properties on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting") 

For Each objItem in colItems
    objItem.EnableReverseDnsLookup = False
    objItem.HopCount = 10
    objItem.MasqueradeDomain = "fabrikam.com"
    objItem.SmartHost = "smtp-server.fabrikam.com"
    objItem.SmartHostType = 1
    objItem.Put_
Next

