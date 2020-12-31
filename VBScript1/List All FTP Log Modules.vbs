' Description: Returns the names of all the FTP log modules on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpInfoSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Log Module List: " & objItem.LogModuleList
Next

