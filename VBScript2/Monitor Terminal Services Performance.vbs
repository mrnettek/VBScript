' Description: Uses cooked performance counters to monitor Terminal Services performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_TermService_TerminalServices").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Active Sessions: " & objItem.ActiveSessions
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Inactive Sessions: " & objItem.InactiveSessions
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Total Sessions: " & objItem.TotalSessions
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

