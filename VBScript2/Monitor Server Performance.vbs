' Description: Uses cooked performance counters to monitor communications using the WINS Server service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfNet_Server").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Blocking Requests Rejected: " & _
            objItem.BlockingRequestsRejected
        Wscript.Echo "Bytes Received Per Second: " & _
            objItem.BytesReceivedPersec
        Wscript.Echo "Bytes Total Per Second: " & objItem.BytesTotalPersec
        Wscript.Echo "Bytes Transmitted Per Second: " & _
            objItem.BytesTransmittedPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Context Blocks Queued Per Second: " & _
            objItem.ContextBlocksQueuedPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Errors Access Permissions: " & _
            objItem.ErrorsAccessPermissions
        Wscript.Echo "Errors Granted Access: " & _
            objItem.ErrorsGrantedAccess
        Wscript.Echo "Errors Logon: " & objItem.ErrorsLogon
        Wscript.Echo "Errors System: " & objItem.ErrorsSystem
        Wscript.Echo "File Directory Searches: " & _
            objItem.FileDirectorySearches
        Wscript.Echo "Files Open: " & objItem.FilesOpen
        Wscript.Echo "Files Opened Total: " & objItem.FilesOpenedTotal
        Wscript.Echo "Logon Per Second: " & objItem.LogonPersec
        Wscript.Echo "Logon Total: " & objItem.LogonTotal
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Pool Nonpaged Bytes: " & objItem.PoolNonpagedBytes
        Wscript.Echo "Pool Nonpaged Failures: " & _
            objItem.PoolNonpagedFailures
        Wscript.Echo "Pool Nonpaged Peak: " & objItem.PoolNonpagedPeak
        Wscript.Echo "Pool Paged Bytes: " & objItem.PoolPagedBytes
        Wscript.Echo "Pool Paged Failures: " & objItem.PoolPagedFailures
        Wscript.Echo "Pool Paged Peak: " & objItem.PoolPagedPeak
        Wscript.Echo "Server Sessions: " & objItem.ServerSessions
        Wscript.Echo "Sessions Errored Out: " & objItem.SessionsErroredOut
        Wscript.Echo "Sessions Forced Off: " & objItem.SessionsForcedOff
        Wscript.Echo "Sessions Logged Off: " & objItem.SessionsLoggedOff
        Wscript.Echo "Sessions Timed Out: " & objItem.SessionsTimedOut
        Wscript.Echo "Work Item Shortages: " & objItem.WorkItemShortages
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

