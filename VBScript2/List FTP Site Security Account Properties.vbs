' Description: Displays default metabase FTP security account settings for an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Allow Anonymous: " & objItem.AllowAnonymous
    Wscript.Echo "Anonymous Only: " & objItem.AnonymousOnly
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & _
        objItem.AnonymousUserPass
Next

