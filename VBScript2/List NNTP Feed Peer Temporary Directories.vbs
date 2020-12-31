' Description: Lists all the NNTP feed peer temporary directories on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpFeedsSetting")
 
For Each objItem in colItems
    Wscript.Echo "Feed Peer Temporary Directory: " & _
        objItem.FeedPeerTempDirectory
    Wscript.Echo "Name: " & objItem.Name
Next

