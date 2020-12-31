' Description: Returns the state of the IMAP server service on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsImapServer")

For Each objItem in colItems
    Wscript.Echo "IMAP Service Version: " & objItem.ImapServiceVersion
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Server State: " & objItem.ServerState
Next

