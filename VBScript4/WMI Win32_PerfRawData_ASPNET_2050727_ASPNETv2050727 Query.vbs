On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_ASPNET_2050727_ASPNETv2050727",,48)
For Each objItem in colItems
    Wscript.Echo "ApplicationRestarts: " & objItem.ApplicationRestarts
    Wscript.Echo "ApplicationsRunning: " & objItem.ApplicationsRunning
    Wscript.Echo "AuditFailureEventsRaised: " & objItem.AuditFailureEventsRaised
    Wscript.Echo "AuditSuccessEventsRaised: " & objItem.AuditSuccessEventsRaised
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ErrorEventsRaised: " & objItem.ErrorEventsRaised
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "InfrastructureErrorEventsRaised: " & objItem.InfrastructureErrorEventsRaised
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "RequestErrorEventsRaised: " & objItem.RequestErrorEventsRaised
    Wscript.Echo "RequestExecutionTime: " & objItem.RequestExecutionTime
    Wscript.Echo "RequestsCurrent: " & objItem.RequestsCurrent
    Wscript.Echo "RequestsDisconnected: " & objItem.RequestsDisconnected
    Wscript.Echo "RequestsQueued: " & objItem.RequestsQueued
    Wscript.Echo "RequestsRejected: " & objItem.RequestsRejected
    Wscript.Echo "RequestWaitTime: " & objItem.RequestWaitTime
    Wscript.Echo "StateServerSessionsAbandoned: " & objItem.StateServerSessionsAbandoned
    Wscript.Echo "StateServerSessionsActive: " & objItem.StateServerSessionsActive
    Wscript.Echo "StateServerSessionsTimedOut: " & objItem.StateServerSessionsTimedOut
    Wscript.Echo "StateServerSessionsTotal: " & objItem.StateServerSessionsTotal
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "WorkerProcessesRunning: " & objItem.WorkerProcessesRunning
    Wscript.Echo "WorkerProcessRestarts: " & objItem.WorkerProcessRestarts
Next

