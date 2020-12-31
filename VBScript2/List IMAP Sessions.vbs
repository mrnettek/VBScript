' Description: Lists the names of all the IMAP sessions running on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsImapSessions")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

