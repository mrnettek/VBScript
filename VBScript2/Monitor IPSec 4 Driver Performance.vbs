' Description: Uses cooked performance counters to monitor IPSec v4 driver performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_IPSec4Driver").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Pending key exchange operations: " & _
            objItem.PendingKeyExchangeOperations
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

