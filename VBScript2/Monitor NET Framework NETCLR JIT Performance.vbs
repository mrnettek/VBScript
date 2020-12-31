' Description: Uses cooked performance counters to monitor CLR JIT performance on a computer running .NET Frameworks 1.1.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_NETFramework_NETCLRJit").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Percent time in JIT: " & objItem.PercentTimeInJIT
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

