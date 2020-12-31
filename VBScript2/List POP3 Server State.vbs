' Description: Lists the state of the POP3 service on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3Server")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    For Each strSource in objItem.Pop3RoutingSources
        Wscript.Echo "Pop3 Routing Sources: " & strSource
    Next
    Wscript.Echo "Pop3 Service Version: " & _
        objItem.Pop3ServiceVersion
    Wscript.Echo "Server State: " & objItem.ServerState
Next

