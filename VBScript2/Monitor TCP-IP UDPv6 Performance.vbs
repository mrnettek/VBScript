' Description: Uses cooked performance counters to measure rates at which UDP datagrams are sent and received by using the UDPv6 protocol.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_UDPv6").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Datagrams sent per second: " & _
            objItem.DatagramsSentPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

