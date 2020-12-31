' Description: Demonstration script that modifies basic FTP site global metabase on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServiceSetting")

For Each objItem in colItems
    objItem.ConnectionTimeout = 600
    objItem.DontLog = True
    objItem.MaxConnections = 50
    objItem.ServerComment = "This server is for IT use only."
    objItem.Put_
Next

