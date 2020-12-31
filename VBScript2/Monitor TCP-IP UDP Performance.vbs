' Description: Uses cooked performance counters to monitor the rates at which UDP datagrams are sent and received by using the UDP protocol.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_UDP").objectSet
objRefresher.Refresh
For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Datagrams No Port Per Second: " & _
            objItem.DatagramsNoPortPersec
        Wscript.Echo "Datagrams Per Second: " & objItem.DatagramsPersec
        Wscript.Echo "Datagrams Received Errors: " & _
            objItem.DatagramsReceivedErrors
        Wscript.Echo "Datagrams Received Per Second: " & _
            objItem.DatagramsReceivedPersec
        Wscript.Echo "Datagrams Sent Per Second: " & _
            objItem.DatagramsSentPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

