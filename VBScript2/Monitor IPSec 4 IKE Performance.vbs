' Description: Uses cooked performance counters to monitor IPSec v4 IKE performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_IPSec4Driver").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Acquire heap size: " & _
            objItem.AcquireHeapSize
        Wscript.Echo "Main mode SA list entries: " & _
            objItem.MainModeSAListEntries
        Wscript.Echo "Quick mode SA list entries: " & _
            objItem.QuickModeSAListEntries
        Wscript.Echo "Receive heap size: " & _
            objItem.ReceiveHeapSize
        Wscript.Echo "Total authentication failures: " & _
            objItem.TotalAuthenticationFailures
        Wscript.Echo "Total main mode SAs: " & _
            objItem.TotalMainModeSAs
        Wscript.Echo "Total negotiation failures: " & _
            objItem.TotalNegotiationFailures
        Wscript.Echo "Total quick mode SAs: " & _
            objItem.TotalQuickModeSAs
        Wscript.Echo "Total soft associations: " & _
            objItem.TotalSoftAssociations
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

