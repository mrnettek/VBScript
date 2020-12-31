' Description: Uses cooked performance counters to measure flow statistics from the packet scheduler


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PSched_PSchedFlow").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Average Packets in Netcard: " & _
            objItem.AveragePacketsinnetcard
        Wscript.Echo "Average Packets in Sequencer: " & _
            objItem.Averagepacketsinsequencer
        Wscript.Echo "Average Packets in Shaper: " & _
            objItem.Averagepacketsinshaper
        Wscript.Echo "Bytes Scheduled: " & objItem.Bytesscheduled
        Wscript.Echo "Bytes Scheduled Per Second: " & _
            objItem.BytesscheduledPersec
        Wscript.Echo "Bytes Transmitted: " & objItem.Bytestransmitted
        Wscript.Echo "Bytes Transmitted Per Second: " & _
            objItem.BytestransmittedPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Maximum Packets in Netcard: " & _
            objItem.MaximumPacketsinnetcard
        Wscript.Echo "Maximum Packets in Sequencer: " & _
            objItem.Maxpacketsinsequencer
        Wscript.Echo "Maximum Packets in Shaper: " & _
            objItem.Maxpacketsinshaper
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Non-conforming Packets Scheduled: " & _
            objItem.Nonconformingpacketsscheduled
        Wscript.Echo "Non-conforming Packets ScheduledPersec: " & _
            objItem.NonconformingpacketsscheduledPersec
        Wscript.Echo "Non-conforming Packets Transmitted: " & _
            objItem.Nonconformingpacketstransmitted
        Wscript.Echo "Non-conforming Packets TransmittedPersec: " & _
            objItem.NonconformingpacketstransmittedPersec
        Wscript.Echo "Packets Dropped: " & objItem.Packetsdropped
        Wscript.Echo "Packets Dropped Per Second: " & _
            objItem.PacketsdroppedPersec
        Wscript.Echo "Packets Scheduled: " & objItem.Packetsscheduled
        Wscript.Echo "Packets Scheduled Per Second: " & _
            objItem.PacketsscheduledPersec
        Wscript.Echo "Packets Transmitted: " & objItem.Packetstransmitted
        Wscript.Echo "Packets Transmitted Per Second: " & _
            objItem.PacketstransmittedPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

