' Description: Uses cooked performance counters to monitor the length of the queues and objects in the queues


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_PerfNet_ServerWorkQueues").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Active Threads: " & objItem.ActiveThreads
        Wscript.Echo "Available Threads: " & objItem.AvailableThreads
        Wscript.Echo "Available Work Items: " & objItem.AvailableWorkItems
        Wscript.Echo "Borrowed Work Items: " & objItem.BorrowedWorkItems
        Wscript.Echo "Bytes Received Per Second: " & _
            objItem.BytesReceivedPersec
        Wscript.Echo "Bytes Sent Per Second: " & objItem.BytesSentPersec
        Wscript.Echo "Bytes Transferred Per Second: " & _
            objItem.BytesTransferredPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Context Blocks Queued Per Second: " & _
            objItem.ContextBlocksQueuedPersec
        Wscript.Echo "Current Clients: " & objItem.CurrentClients
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Queue Length: " & objItem.QueueLength
        Wscript.Echo "Read Bytes Per Second: " & objItem.ReadBytesPersec
        Wscript.Echo "Read Operations Per Second: " & _
            objItem.ReadOperationsPersec
        Wscript.Echo "Total Bytes Per Second: " & objItem.TotalBytesPersec
        Wscript.Echo "Total Operations Per Second: " & _
            objItem.TotalOperationsPersec
        Wscript.Echo "Work Item Shortages: " & objItem.WorkItemShortages
        Wscript.Echo "Write Bytes Per Second: " & objItem.WriteBytesPersec
        Wscript.Echo "Write Operations Per Second: " & _
            objItem.WriteOperationsPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

