' Description: Returns the state of the NNTP server service on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServer")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NNTP Service Version: " &  _
        objItem.NntpServiceVersion
    Wscript.Echo "Server State: " & objItem.ServerState
Next

