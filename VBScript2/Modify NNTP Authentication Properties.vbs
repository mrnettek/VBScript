' Description: Demonstration script that modifies global NNTP authentication metabase property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting") 

For Each objItem in colItems
    objItem.AccessSSLMapCert = False
    objItem.AccessSSLNegotiateCert = False
    objItem.AccessSSLRequireCert = False
    objItem.AuthAnonymous = True
    objItem.AuthBasic = True
    objItem.AuthNTLM = False
    objItem.Put_
Next

