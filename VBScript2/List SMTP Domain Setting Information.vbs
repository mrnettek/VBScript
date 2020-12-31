' Description: Displays SMTP domain setting information on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpDomainSetting")
 
For Each objItem in colItems
    For Each strTurn in objItem.AuthTurnList
        Wscript.Echo "Authentication Turn List: " & strTurn
    Next
    Wscript.Echo "CSide Etrn Domains: " & objItem.CSideEtrnDomains
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Relay For Authentication: " & objItem.RelayForAuth
    Wscript.Echo "Relay IP List: " & objItem.RelayIpList
    Wscript.Echo "Route Action: " & objItem.RouteAction
    Wscript.Echo "Route Action String: " & objItem.RouteActionString
    Wscript.Echo "Route Password: " & objItem.RoutePassword
    Wscript.Echo "Route User Name: " & objItem.RouteUserName
Next

