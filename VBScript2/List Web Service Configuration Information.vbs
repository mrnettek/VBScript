' Description: Displays information about the IIS Web service service on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebService")
 
For Each objItem in colItems
    Wscript.Echo "Accept Pause: " & objItem.AcceptPause
    Wscript.Echo "Accept Stop: " & objItem.AcceptStop
    Wscript.Echo "Application Isolated: " & objItem.AppIsolated
    Wscript.Echo "Application Package ID: " & objItem.AppPackageID
    Wscript.Echo "Application Package Name: " & _
        objItem.AppPackageName
    Wscript.Echo "Application Root: " & objItem.AppRoot
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Check Point: " & objItem.CheckPoint
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Desktop Interact: " & objItem.DesktopInteract
    Wscript.Echo "Display Name: " & objItem.DisplayName
    Wscript.Echo "Error Control: " & objItem.ErrorControl
    Wscript.Echo "Exit Code: " & objItem.ExitCode
    Wscript.Echo "Installation Date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path Name: " & objItem.PathName
    Wscript.Echo "Process ID: " & objItem.ProcessId
    Wscript.Echo "Service Specific Exit Code: " & _
        objItem.ServiceSpecificExitCode
    Wscript.Echo "Service Type: " & objItem.ServiceType
    Wscript.Echo "SSL Certificate Hash: " & objItem.SSLCertHash
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "Start Mode: " & objItem.StartMode
    Wscript.Echo "Start Name: " & objItem.StartName
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag Id: " & objItem.TagId
    Wscript.Echo "Wait Hint: " & objItem.WaitHint
Next

