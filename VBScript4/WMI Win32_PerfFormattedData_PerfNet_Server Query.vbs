On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfNet_Server",,48)
For Each objItem in colItems
    Wscript.Echo "BlockingRequestsRejected: " & objItem.BlockingRequestsRejected
    Wscript.Echo "BytesReceivedPersec: " & objItem.BytesReceivedPersec
    Wscript.Echo "BytesTotalPersec: " & objItem.BytesTotalPersec
    Wscript.Echo "BytesTransmittedPersec: " & objItem.BytesTransmittedPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContextBlocksQueuedPersec: " & objItem.ContextBlocksQueuedPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ErrorsAccessPermissions: " & objItem.ErrorsAccessPermissions
    Wscript.Echo "ErrorsGrantedAccess: " & objItem.ErrorsGrantedAccess
    Wscript.Echo "ErrorsLogon: " & objItem.ErrorsLogon
    Wscript.Echo "ErrorsSystem: " & objItem.ErrorsSystem
    Wscript.Echo "FileDirectorySearches: " & objItem.FileDirectorySearches
    Wscript.Echo "FilesOpen: " & objItem.FilesOpen
    Wscript.Echo "FilesOpenedTotal: " & objItem.FilesOpenedTotal
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "LogonPersec: " & objItem.LogonPersec
    Wscript.Echo "LogonTotal: " & objItem.LogonTotal
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PoolNonpagedBytes: " & objItem.PoolNonpagedBytes
    Wscript.Echo "PoolNonpagedFailures: " & objItem.PoolNonpagedFailures
    Wscript.Echo "PoolNonpagedPeak: " & objItem.PoolNonpagedPeak
    Wscript.Echo "PoolPagedBytes: " & objItem.PoolPagedBytes
    Wscript.Echo "PoolPagedFailures: " & objItem.PoolPagedFailures
    Wscript.Echo "PoolPagedPeak: " & objItem.PoolPagedPeak
    Wscript.Echo "ServerSessions: " & objItem.ServerSessions
    Wscript.Echo "SessionsErroredOut: " & objItem.SessionsErroredOut
    Wscript.Echo "SessionsForcedOff: " & objItem.SessionsForcedOff
    Wscript.Echo "SessionsLoggedOff: " & objItem.SessionsLoggedOff
    Wscript.Echo "SessionsTimedOut: " & objItem.SessionsTimedOut
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "WorkItemShortages: " & objItem.WorkItemShortages
Next

