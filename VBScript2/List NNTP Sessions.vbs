' Description: Lists the names of all the NNTP sessions on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsNntpSessions")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

