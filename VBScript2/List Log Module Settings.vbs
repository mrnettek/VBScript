' Description: Returns information about all the log module settings on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsLogModuleSetting")
 
For Each objItem in colItems
    Wscript.Echo "Log Module Id: " & objItem.LogModuleId
    Wscript.Echo "Log Module UI Id: " & objItem.LogModuleUiId
    Wscript.Echo "Name: " & objItem.Name
Next

