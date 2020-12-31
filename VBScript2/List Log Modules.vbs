' Description: Lists the names of all the log modules found on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsLogModules")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

