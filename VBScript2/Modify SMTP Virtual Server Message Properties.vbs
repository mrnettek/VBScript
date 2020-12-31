' Description: Demonstration script that modifies global SMTP virtual server message metabase property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

 For Each objItem in colItems
    objItem.BadMailDirectory = "C:\Badmail"
    objItem.MaxBatchedMessages = 30
    objItem.MaxMessageSize = 500000
    objItem.MaxRecipients = 50
    objItem.MaxSessionSize = 1000000
    objItem.SendNdrTo = "email-admin@fabrikam.com"
    objItem.Put_
Next

