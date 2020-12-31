' Description: Demonstration script that modifies default Web site performance metabase properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    objItem.MaxBandwidth = -1
    objItem.MaxConnections = 10000
    objItem.Put_
Next

