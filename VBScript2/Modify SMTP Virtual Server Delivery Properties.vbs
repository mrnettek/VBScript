' Description: Demonstration script that modifies global SMTP virtual directory metabase property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    objItem.SmtpLocalDelayExpireMinutes = 1000
    objItem.SmtpLocalNDRExpireMinutes = 4000
    objItem.SmtpRemoteDelayExpireMinutes = 1000
    objItem.SmtpRemoteNDRExpireMinutes = 4000
    objItem.SmtpRemoteProgressiveRetry = "60,120,240"
    objItem.Put_
Next

