' Description: Returns information about all the Web directories on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebDirectory")
 
For Each objItem in colItems
    Wscript.Echo "Application Isolated: " & objItem.AppIsolated
    Wscript.Echo "Application Package ID: " & objItem.AppPackageID
    Wscript.Echo "Application Package Name: " & objItem.AppPackageName
    Wscript.Echo "Application Root: " & objItem.AppRoot
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Status: " & objItem.Status
Next

