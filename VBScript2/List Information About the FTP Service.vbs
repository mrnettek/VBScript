' Description: Lists configuration settings for the FTP server service on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsFtpService")

For Each objItem in colItems
    Wscript.Echo "Accept Pause: " & objItem.AcceptPause
    Wscript.Echo "Accept Stop: " & objItem.AcceptStop
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
    Wscript.Echo "Process Id: " & objItem.ProcessId
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

