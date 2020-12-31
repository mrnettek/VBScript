' Description: Uses cooked performance counters to measure pipe statistics from the packet scheduler


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PSched_PSchedPipe").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Average Packets in Netcard: " & _
            objItem.Averagepacketsinnetcard
        Wscript.Echo "Average Packets in Sequencer: " & _
            objItem.Averagepacketsinsequencer
        Wscript.Echo "Average Packets in Shaper: " & _
            objItem.Averagepacketsinshaper
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Flow Mods Rejected: " & objItem.Flowmodsrejected
        Wscript.Echo "Flows Closed: " & objItem.Flowsclosed
        Wscript.Echo "Flows Modified: " & objItem.Flowsmodified
        Wscript.Echo "Flows Opened: " & objItem.Flowsopened
        Wscript.Echo "Flows Rejected: " & objItem.Flowsrejected
        Wscript.Echo "Maximum Packets in Netcard: " & _
            objItem.Maxpacketsinnetcard
        Wscript.Echo "Maximum Packets in Sequencer: " & _
            objItem.Maxpacketsinsequencer
        Wscript.Echo "Maximum Packets in Shaper: " & _
            objItem.Maxpacketsinshaper
        Wscript.Echo "Maximum Simultaneous Flows: " & _
            objItem.Maxsimultaneousflows
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Non-conforming Packets Scheduled: " & _
            objItem.Nonconformingpacketsscheduled
        Wscript.Echo "Non-conforming Packets ScheduledPersec: " & _
            objItem.NonconformingpacketsscheduledPersec
        Wscript.Echo "Non-conforming Packets Transmitted: " & _
            objItem.Nonconformingpacketstransmitted
        Wscript.Echo "Non-conforming Packets TransmittedPersec: " & _
            objItem.NonconformingpacketstransmittedPersec
        Wscript.Echo "Out-of-packets: " & objItem.Outofpackets
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

