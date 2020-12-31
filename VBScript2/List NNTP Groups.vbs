' Description: Lists all the NNTP groups on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsNntpGroups")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

