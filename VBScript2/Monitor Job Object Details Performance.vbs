' Description: Uses cooked performance counters to return detailed performance information about the processes that make up a job object.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_PerfProc_JobObjectDetails").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Creating Process ID: " & objItem.CreatingProcessID
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Elapsed Time: " & objItem.ElapsedTime
        Wscript.Echo "Handle Count: " & objItem.HandleCount
        Wscript.Echo "ID Process: " & objItem.IDProcess
        Wscript.Echo "I/O Data Bytes Per Second: " & objItem.IODataBytesPersec
        Wscript.Echo "I/O Data Operations Per Second: " & _
            objItem.IODataOperationsPersec
        Wscript.Echo "I/O Other Bytes Per Second: " & _
            objItem.IOOtherBytesPersec
        Wscript.Echo "I/O Other Operations Per Second: " & _
            objItem.IOOtherOperationsPersec
        Wscript.Echo "I/O Read Bytes Per Second: " & objItem.IOReadBytesPersec
        Wscript.Echo "I/O Read Operations Per Second: " & _
            objItem.IOReadOperationsPersec
        Wscript.Echo "I/O Write Bytes Per Second: " & _
            objItem.IOWriteBytesPersec
        Wscript.Echo "I/O Write Operations Per Second: " & _
            objItem.IOWriteOperationsPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Page Faults Per Second: " & objItem.PageFaultsPersec
        Wscript.Echo "Page File Bytes: " & objItem.PageFileBytes
        Wscript.Echo "Page File Bytes Peak: " & objItem.PageFileBytesPeak
        Wscript.Echo "Percent Privileged Time: " & _
            objItem.PercentPrivilegedTime
        Wscript.Echo "Percent Processor Time: " & objItem.PercentProcessorTime
        Wscript.Echo "Percent User Time: " & objItem.PercentUserTime
        Wscript.Echo "Pool Nonpaged Bytes: " & objItem.PoolNonpagedBytes
        Wscript.Echo "Pool Paged Bytes: " & objItem.PoolPagedBytes
        Wscript.Echo "Priority Base: " & objItem.PriorityBase
        Wscript.Echo "Private Bytes: " & objItem.PrivateBytes
        Wscript.Echo "Thread Count: " & objItem.ThreadCount
        Wscript.Echo "Virtual Bytes: " & objItem.VirtualBytes
        Wscript.Echo "Virtual Bytes Peak: " & objItem.VirtualBytesPeak
        Wscript.Echo "Working Set: " & objItem.WorkingSet
        Wscript.Echo "Working Set Peak: " & objItem.WorkingSetPeak
        Wscript.Sleep 2000
    objRefresher.Refresh
    Next
Next

