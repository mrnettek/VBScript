' Description: Returns global Web server information for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebServer")
 
For Each objItem in colItems
    Wscript.Echo "Application Isolated: " & objItem.AppIsolated
    Wscript.Echo "Application Package ID: " & objItem.AppPackageID
    Wscript.Echo "Application Package Name: " & objItem.AppPackageName
    Wscript.Echo "Application Root: " & objItem.AppRoot
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Installation Date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Server State: " & objItem.ServerState
    Wscript.Echo "Status: " & objItem.Status
Next

