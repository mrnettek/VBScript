' Description: Uses cooked performance counters to return page file performance information.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_PagingFile").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Usage: " & objItem.PercentUsage
        Wscript.Echo "Percent Usage Peak: " & objItem.PercentUsagePeak
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

