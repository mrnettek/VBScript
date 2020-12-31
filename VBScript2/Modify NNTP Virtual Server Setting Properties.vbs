' Description: Demonstration script that modifies global NNTP virtual server property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting") 

For Each objItem in colItems
    objItem.AdminEmail = "admin@fabrikam.com"
    objItem.AllowClientPosts = True
    objItem.AllowControlMsgs = True
    objItem.AllowFeedPosts = True
    objItem.ClientPostHardLimit = 1000000
    objItem.ClientPostSoftLimit = 250000
    objItem.DefaultModeratorDomain = "fabrikam.com"
    objItem.FeedPostHardLimit = 1000000
    objItem.FeedPostSoftLimit = 250000
    objItem.SmtpServer = "mail.fabrikam.com"
    objItem.Put_
Next

