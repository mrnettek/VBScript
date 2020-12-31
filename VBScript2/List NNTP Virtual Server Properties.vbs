' Description: Returns global NNTP connection metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Connection Timeout: " & objItem.ConnectionTimeout
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "NNTP UUCP Name: " & objItem.NntpUucpName
Next

