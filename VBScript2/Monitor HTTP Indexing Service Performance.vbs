' Description: Uses cooked performance counters to return performance data from the HTTP Indexing Service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService," & _
    "Win32_PerfFormattedData_ISAPISearch_HTTPIndexingService").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Active Queries: " & objItem.Activequeries
        Wscript.Echo "Cache Items: " & objItem.Cacheitems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Current Requests Queued: " & _
            objItem.Currentrequestsqueued
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Cache Hits: " & objItem.PercentCachehits
        Wscript.Echo "Percent Cache Misses: " & objItem.PercentCachemisses
        Wscript.Echo "Queries Per Minute: " & objItem.Queriesperminute
        Wscript.Echo "Total Queries: " & objItem.Totalqueries
        Wscript.Echo "Total Requests Rejected: " & _
            objItem.Totalrequestsrejected
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

