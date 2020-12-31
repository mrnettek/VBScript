' Description: Uses cooked performance counters to monitor counters that apply to more than one instance of a component processors on the computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_System").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Alignment Fixups Per Second: " & _
            objItem.AlignmentFixupsPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Context Switches Per Second: " & _
            objItem.ContextSwitchesPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Exception Dispatches Per Second: " & _
            objItem.ExceptionDispatchesPersec
        Wscript.Echo "File Control Bytes Per Second: " & _
            objItem.FileControlBytesPersec
        Wscript.Echo "File Control Operations Per Second: " & _
            objItem.FileControlOperationsPersec
        Wscript.Echo "File Data Operations Per Second: " & _
            objItem.FileDataOperationsPersec
        Wscript.Echo "File Read Bytes Per Second: " & _
            objItem.FileReadBytesPersec
        Wscript.Echo "File Read Operations Per Second: " & _
            objItem.FileReadOperationsPersec
        Wscript.Echo "File Write Bytes Per Second: " & _
            objItem.FileWriteBytesPersec
        Wscript.Echo "File Write Operations Per Second: " & _
            objItem.FileWriteOperationsPersec
        Wscript.Echo "Floating Emulations Per Second: " & _
            objItem.FloatingEmulationsPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Registry Quota In Use: " & _
            objItem.PercentRegistryQuotaInUse
        Wscript.Echo "Processes: " & objItem.Processes
        Wscript.Echo "Processor Queue Length: " & _
            objItem.ProcessorQueueLength
        Wscript.Echo "System Calls Per Second: " & _
            objItem.SystemCallsPersec
        Wscript.Echo "System UpTime: " & objItem.SystemUpTime
        Wscript.Echo "Threads: " & objItem.Threads
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

