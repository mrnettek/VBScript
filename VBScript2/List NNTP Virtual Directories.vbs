' Description: Lists the names of all the NNTP virtual directories on a server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsNntpVirtualDir")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

