' Description: Uses cooked performance counters to return information about accounting and processor usage data collected by each active named job object.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfProc_JobObject").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Current Percent Kernel Mode Time: " & _
            objItem.CurrentPercentKernelModeTime
        Wscript.Echo "Current Percent Processor Time: " & _
            objItem.CurrentPercentProcessorTime
        Wscript.Echo "Current Percent User Mode Time: " & _\
            objItem.CurrentPercentUserModeTime
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Pages Per Second: " & objItem.PagesPerSec
        Wscript.Echo "Process Count Active: " & objItem.ProcessCountActive
        Wscript.Echo "Process Count Terminated: " & _
            objItem.ProcessCountTerminated
        Wscript.Echo "Process Count Total: " & objItem.ProcessCountTotal
        Wscript.Echo "This Period Milliseconds Kernel Mode: " & _
            objItem.ThisPeriodmSecKernelMode
        Wscript.Echo "This Period Milliseconds Processor: " & _
            objItem.ThisPeriodmSecProcessor
        Wscript.Echo "This Period Milliseconds User Mode: " & _
            objItem.ThisPeriodmSecUserMode
        Wscript.Echo "Total Milliseconds Kernel Mode: " & _
            objItem.TotalmSecKernelMode
        Wscript.Echo "Total Milliseconds Processor: " & _
            objItem.TotalmSecProcessor
        Wscript.Echo "Total Milliseconds User Mode: " & _
            objItem.TotalmSecUserMode
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

