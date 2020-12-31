' Description: Demonstration script that modifies global FTP security account metabase settings on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServiceSetting")

For Each objItem in colItems
    objItem.AllowAnonymous = True
    objItem.AnonymousOnly = True
    objItem.AnonymousUserName = "TestUser"
    objItem.AnonymousUserPass = "password"
    objItem.Put_
Next

