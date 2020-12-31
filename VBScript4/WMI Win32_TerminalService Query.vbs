On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TerminalService",,48)
For Each objItem in colItems
    Wscript.Echo "AcceptPause: " & objItem.AcceptPause
    Wscript.Echo "AcceptStop: " & objItem.AcceptStop
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckPoint: " & objItem.CheckPoint
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DesktopInteract: " & objItem.DesktopInteract
    Wscript.Echo "DisconnectedSessions: " & objItem.DisconnectedSessions
    Wscript.Echo "DisplayName: " & objItem.DisplayName
    Wscript.Echo "ErrorControl: " & objItem.ErrorControl
    Wscript.Echo "EstimatedSessionCapacity: " & objItem.EstimatedSessionCapacity
    Wscript.Echo "ExitCode: " & objItem.ExitCode
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PathName: " & objItem.PathName
    Wscript.Echo "ProcessId: " & objItem.ProcessId
    Wscript.Echo "RawSessionCapacity: " & objItem.RawSessionCapacity
    Wscript.Echo "ResourceConstraint: " & objItem.ResourceConstraint
    Wscript.Echo "ServiceSpecificExitCode: " & objItem.ServiceSpecificExitCode
    Wscript.Echo "ServiceType: " & objItem.ServiceType
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "StartMode: " & objItem.StartMode
    Wscript.Echo "StartName: " & objItem.StartName
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "TagId: " & objItem.TagId
    Wscript.Echo "TotalSessions: " & objItem.TotalSessions
    Wscript.Echo "WaitHint: " & objItem.WaitHint
Next

