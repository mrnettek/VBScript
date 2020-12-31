' Description: Returns global NNTP virtual server metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting") 

For Each objItem in colItems
    Wscript.Echo "Admin Email: " & objItem.AdminEmail
    Wscript.Echo "Allow Client Posts: " & objItem.AllowClientPosts
    Wscript.Echo "Allow Control Msgs: " & objItem.AllowControlMsgs
    Wscript.Echo "Allow Feed Posts: " & objItem.AllowFeedPosts
    Wscript.Echo "Client Post Hard Limit: " & _
        objItem.ClientPostHardLimit
    Wscript.Echo "Client Post Soft Limit: " & _
        objItem.ClientPostSoftLimit
    Wscript.Echo "Default Moderator Domain: " & _
        objItem.DefaultModeratorDomain
    Wscript.Echo "Feed Post Hard Limit: " & objItem.FeedPostHardLimit
    Wscript.Echo "Feed Post Soft Limit: " & objItem.FeedPostSoftLimit
    Wscript.Echo "SMTP Server: " & objItem.SmtpServer
Next

