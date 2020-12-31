' Description: Uses cooked performance counters to monitor the rates at which IP datagrams are sent and received by using IPv4 protocols.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_IPv4").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Datagrams per second: " & objItem.DatagramsPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

