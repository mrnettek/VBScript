' Description: Uses cooked performance counters to monitor CLR locks and threads performance on a computer running .NET Frameworks 1.1.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_NETFramework_NETCLRLocksAndThreads").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Current queue length: " & objItem.CurrentQueueLength
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

