' Description: Uses cooked performance counters to monitor the rates at which bytes and packets are sent and received over a TCP/IP network connection.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_NetworkInterface").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Bytes Received Per Second: " & _
        objItem.BytesReceivedPersec
        Wscript.Echo "Bytes Sent Per Second: " & objItem.BytesSentPersec
        Wscript.Echo "Bytes Total Per Second: " & objItem.BytesTotalPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Current Bandwidth: " & objItem.CurrentBandwidth
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Output Queue Length: " & objItem.OutputQueueLength
        Wscript.Echo "Packets Outbound Discarded: " & _
            objItem.PacketsOutboundDiscarded
        Wscript.Echo "Packets Outbound Errors: " & _
            objItem.PacketsOutboundErrors
        Wscript.Echo "Packets Per Second: " & objItem.PacketsPersec
        Wscript.Echo "Packets Received Discarded: " & _
            objItem.PacketsReceivedDiscarded
        Wscript.Echo "Packets Received Errors: " & _
            objItem.PacketsReceivedErrors
        Wscript.Echo "Packets Received Non-Unicast Per Second: " & _
            objItem.PacketsReceivedNonUnicastPersec
        Wscript.Echo "Packets Received Per Second: " & _
            objItem.PacketsReceivedPersec
        Wscript.Echo "Packets Received Unicast Per Second: " & _
            objItem.PacketsReceivedUnicastPersec
        Wscript.Echo "Packets Received Unknown: " & _
            objItem.PacketsReceivedUnknown
        Wscript.Echo "Packets Sent Non-Unicast Per Second: " & _
            objItem.PacketsSentNonUnicastPersec
        Wscript.Echo "Packets Sent Per Second: " & objItem.PacketsSentPersec
        Wscript.Echo "Packets Sent Unicast Per Second: " & _
            objItem.PacketsSentUnicastPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

