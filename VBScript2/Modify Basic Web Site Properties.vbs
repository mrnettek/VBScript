' Description: Demonstration script that modifies default Web site metabase properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    objItem.AllowKeepAlive = True
    objItem.ConnectionTimeout = 1200
    objItem.DontLog = False
    objItem.ServerComment = "This is an intranet-only server."
    objItem.Put_
Next

