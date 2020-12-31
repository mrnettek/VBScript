' Description: Returns global SMTP delivery metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

For Each objItem in colItems
    Wscript.Echo "SMTP Local Delay Expire Minutes: " & _
        objItem.SmtpLocalDelayExpireMinutes
    Wscript.Echo "SMTP Local NDR Expire Minutes: " & _
        objItem.SmtpLocalNDRExpireMinutes
    Wscript.Echo "SMTP Remote Delay Expire Minutes: " & _
        objItem.SmtpRemoteDelayExpireMinutes
    Wscript.Echo "SMTP Remote NDR Expire Minutes: " & _
        objItem.SmtpRemoteNDRExpireMinutes
    Wscript.Echo "SMTP Remote Progressive Retry: " & _
        objItem.SmtpRemoteProgressiveRetry
Next

