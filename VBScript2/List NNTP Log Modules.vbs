' Description: Lists all the NNTP log modules on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpInfoSetting")
 
For Each objItem in colItems
    Wscript.Echo "Log Module List: " & objItem.LogModuleList
    Wscript.Echo "Name: " & objItem.Name
Next

