' Description: Demonstration script that modifies global SMTP authentication metabase properties on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    objItem.AuthAnonymous = True
    objItem.AuthBasic = True
    objItem.AuthNTLM = True
    objItem.SaslLogonDomain = "fabrikam.com"
    objItem.Put_
Next

