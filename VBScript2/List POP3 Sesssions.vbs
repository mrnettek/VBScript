' Description: Lists the names of all the POP3 sessions on a server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3Sessions")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

