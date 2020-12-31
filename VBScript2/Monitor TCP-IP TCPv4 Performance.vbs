' Description: Uses cooked performance counters to monitor the rates at which TCP segments are sent and received by using the TCPv4 protocol.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_TCPv4").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Connection failures: " & objItem.ConnectionFailures
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

