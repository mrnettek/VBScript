' Description: Uses cooked performance counters to monitor the rates at which TCP segments are sent and received by using the TCP protocol.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_TCP").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Connection Failures: " & objItem.ConnectionFailures
        Wscript.Echo "Connections Active: " & objItem.ConnectionsActive
        Wscript.Echo "Connections Established: " & _
            objItem.ConnectionsEstablished
        Wscript.Echo "Connections Passive: " & objItem.ConnectionsPassive
        Wscript.Echo "Connections Reset: " & objItem.ConnectionsReset
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Segments Per Second: " & objItem.SegmentsPersec
        Wscript.Echo "Segments Received Per Second: " & _\
            objItem.SegmentsReceivedPersec
        Wscript.Echo "Segments Retransmitted Per Second: " & _
            objItem.SegmentsRetransmittedPersec
        Wscript.Echo "Segments Sent Per Second: " & _
             objItem.SegmentsSentPersec
        objRefresher.Refresh
    Next
Next

