' Description: Displays performance information for all Web sites on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Maximum Bandwidth: " & objItem.MaxBandwidth
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
Next

