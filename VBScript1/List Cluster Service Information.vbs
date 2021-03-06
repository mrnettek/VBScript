' Description: Returns information about the cluster service on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")

Set colItems = objWMIService.ExecQuery("Select * from MSCluster_Service")

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Enable event log replication: " & _
        objItem.EnableEventLogReplication
    Wscript.Echo "Installation date: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Node highest version: " & objItem.NodeHighestVersion
    Wscript.Echo "Node lowest version: " & objItem.NodeLowestVersion
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "Start mode: " & objItem.StartMode
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "System name: " & objItem.SystemName
    Wscript.Echo
Next

