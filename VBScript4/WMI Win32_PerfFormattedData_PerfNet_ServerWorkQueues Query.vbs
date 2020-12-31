On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfNet_ServerWorkQueues",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveThreads: " & objItem.ActiveThreads
    Wscript.Echo "AvailableThreads: " & objItem.AvailableThreads
    Wscript.Echo "AvailableWorkItems: " & objItem.AvailableWorkItems
    Wscript.Echo "BorrowedWorkItems: " & objItem.BorrowedWorkItems
    Wscript.Echo "BytesReceivedPersec: " & objItem.BytesReceivedPersec
    Wscript.Echo "BytesSentPersec: " & objItem.BytesSentPersec
    Wscript.Echo "BytesTransferredPersec: " & objItem.BytesTransferredPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContextBlocksQueuedPersec: " & objItem.ContextBlocksQueuedPersec
    Wscript.Echo "CurrentClients: " & objItem.CurrentClients
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "QueueLength: " & objItem.QueueLength
    Wscript.Echo "ReadBytesPersec: " & objItem.ReadBytesPersec
    Wscript.Echo "ReadOperationsPersec: " & objItem.ReadOperationsPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TotalBytesPersec: " & objItem.TotalBytesPersec
    Wscript.Echo "TotalOperationsPersec: " & objItem.TotalOperationsPersec
    Wscript.Echo "WorkItemShortages: " & objItem.WorkItemShortages
    Wscript.Echo "WriteBytesPersec: " & objItem.WriteBytesPersec
    Wscript.Echo "WriteOperationsPersec: " & objItem.WriteOperationsPersec
Next

