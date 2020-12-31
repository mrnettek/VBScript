' Description: Uses cooked performance counters to monitor the rates at which messages are sent and received by using ICMPv6 protocols.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_ICMPv6").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Messages per second: " & objItem.MessagesPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

