' Description: Uses cooked performance counters to return information about thread behavior.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfProc_Thread").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Context Switches Per Second: " & _
            objItem.ContextSwitchesPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Elapsed Time: " & objItem.ElapsedTime
        Wscript.Echo "ID Process: " & objItem.IDProcess
        Wscript.Echo "ID Thread: " & objItem.IDThread
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Privileged Time: " & _
            objItem.PercentPrivilegedTime
        Wscript.Echo "Percent Processor Time: " & objItem.PercentProcessorTime
        Wscript.Echo "Percent User Time: " & objItem.PercentUserTime
        Wscript.Echo "Priority Base: " & objItem.PriorityBase
        Wscript.Echo "Priority Current: " & objItem.PriorityCurrent
        Wscript.Echo "Start Address: " & objItem.StartAddress
        Wscript.Echo "Thread State: " & objItem.ThreadState
        Wscript.Echo "Thread Wait Reason: " & objItem.ThreadWaitReason
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

