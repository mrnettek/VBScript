' Description: Returns information about the configuration of the POP3 server service on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3Service")

For Each objItem in colItems
    Wscript.Echo "Accept Pause: " & objItem.AcceptPause
    Wscript.Echo "Accept Stop: " & objItem.AcceptStop
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckPoint: " & objItem.CheckPoint
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Desktop Interact: " & _
        objItem.DesktopInteract
    Wscript.Echo "Display Name: " & objItem.DisplayName
    Wscript.Echo "Error Control: " & objItem.ErrorControl
    Wscript.Echo "Exit Code: " & objItem.ExitCode
    Wscript.Echo "Installation Date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path Name: " & objItem.PathName
    For Each strSource in objItem.Pop3RoutingSources
        Wscript.Echo "Pop3 Routing Sources: " & strSource
    Next
    Wscript.Echo "Pop3 Service Version: " & _
        objItem.Pop3ServiceVersion
    Wscript.Echo "ProcessId: " & objItem.ProcessId
    Wscript.Echo "Service Specific Exit Code: " & _
        objItem.ServiceSpecificExitCode
    Wscript.Echo "Service Type: " & objItem.ServiceType
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "Start Mode: " & objItem.StartMode
    Wscript.Echo "Start Name: " & objItem.StartName
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag Id: " & objItem.TagId
    Wscript.Echo "Wait Hint: " & objItem.WaitHint
Next

