' Description: Demonstration script that modifies a Web info metabase property (ServerConfigSSLAllowEncrypt) on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebInfoSetting")
 
For Each objItem in colItems
    objItem.ServerConfigSSLAllowEncrypt = TRUE
    objItem.Put_
Next

