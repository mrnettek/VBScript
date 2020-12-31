' Description: Lists all the SMTP log modules on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpInfoSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Log Module List: " & objItem.LogModuleList
Next

