' Description: Returns global SMTP message metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")

 For Each objItem in colItems
    Wscript.Echo "Bad Mail Directory: " & objItem.BadMailDirectory
    Wscript.Echo "Maximum Batched Messages: " & _
        objItem.MaxBatchedMessages
    Wscript.Echo "Maximum Message Size: " & objItem.MaxMessageSize
    Wscript.Echo "Maximum Recipients: " & objItem.MaxRecipients
    Wscript.Echo "Maximum Session Size: " & objItem.MaxSessionSize
    Wscript.Echo "Send Ndr To: " & objItem.SendNdrTo
Next

