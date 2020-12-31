' Description: Displays default metabase FTP site messages for an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServiceSetting")

For Each objItem in colItems
    For Each objMessage in objItem.BannerMessage
        Wscript.Echo "Banner Message: " & objMessage
    Next
    Wscript.Echo "Exit Message: " & objItem.ExitMessage
    For Each objMessage in objItem.GreetingMessage
        Wscript.Echo "Greeting Message: " & objMessage
    Next
    Wscript.Echo "Maximum Clients Message: " & _
        objItem.MaxClientsMessage
Next

