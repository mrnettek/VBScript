' Description: Uses cooked performance counters to monitor network connections originating at the local computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfNet_Redirector").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Bytes Received Per Second: " & _
            objItem.BytesReceivedPersec
        Wscript.Echo "Bytes Total Per Second: " & objItem.BytesTotalPersec
        Wscript.Echo "Bytes Transmitted Per Second: " & _
            objItem.BytesTransmittedPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Connects Core: " & objItem.ConnectsCore
        Wscript.Echo "Connects Lan Manager 2.0: " & _
            objItem.ConnectsLanManager20
        Wscript.Echo "Connects Lan Manager 2.1: " & _
            objItem.ConnectsLanManager21
        Wscript.Echo "Connects Windows NT: " & objItem.ConnectsWindowsNT
        Wscript.Echo "Current Commands: " & objItem.CurrentCommands
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "File Data Operations Per Second: " & _
            objItem.FileDataOperationsPersec
        Wscript.Echo "File Read Operations Per Second: " & _
            objItem.FileReadOperationsPersec
        Wscript.Echo "File Write Operations Per Second: " & _
            objItem.FileWriteOperationsPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Network Errors Per Second: " & _
            objItem.NetworkErrorsPersec
        Wscript.Echo "Packets Per Second: " & objItem.PacketsPersec
        Wscript.Echo "Packets Received Per Second: " & _
            objItem.PacketsReceivedPersec
        Wscript.Echo "Packets Transmitted Per Second: " & _
            objItem.PacketsTransmittedPersec
        Wscript.Echo "Read Bytes Cache Per Second: " & _
            objItem.ReadBytesCachePersec
        Wscript.Echo "Read Bytes Network Per Second: " & _
            objItem.ReadBytesNetworkPersec
        Wscript.Echo "Read Bytes NonPaging Per Second: " & _
            objItem.ReadBytesNonPagingPersec
        Wscript.Echo "Read Bytes Paging Per Second: " & _
            objItem.ReadBytesPagingPersec
        Wscript.Echo "Read Operations Random Per Second: " & _
            objItem.ReadOperationsRandomPersec
        Wscript.Echo "Read Packets Per Second: " & objItem.ReadPacketsPersec
        Wscript.Echo "Read Packets Small Per Second: " & _
            objItem.ReadPacketsSmallPersec
        Wscript.Echo "Reads Denied Per Second: " & objItem.ReadsDeniedPersec
        Wscript.Echo "Reads Large Per Second: " & objItem.ReadsLargePersec
        Wscript.Echo "Server Disconnects: " & objItem.ServerDisconnects
        Wscript.Echo "Server Reconnects: " & objItem.ServerReconnects
        Wscript.Echo "Server Sessions: " & objItem.ServerSessions
        Wscript.Echo "Server Sessions Hung: " & objItem.ServerSessionsHung
        Wscript.Echo "Write Bytes Cache Per Second: " & _
            objItem.WriteBytesCachePersec
        Wscript.Echo "Write Bytes Network Per Second: " & _
            objItem.WriteBytesNetworkPersec
        Wscript.Echo "Write Bytes NonPaging Per Second: " & _
            objItem.WriteBytesNonPagingPersec
        Wscript.Echo "Write Bytes Paging Per Second: " & _
            objItem.WriteBytesPagingPersec
        Wscript.Echo "Write Operations Random Per Second: " & _
            objItem.WriteOperationsRandomPersec
        Wscript.Echo "Write Packets Per Second: " & _
            objItem.WritePacketsPersec
        Wscript.Echo "Write PacketsSmall Per Second: " & _
            objItem.WritePacketsSmallPersec
        Wscript.Echo "Writes Denied Per Second: " & objItem.WritesDeniedPersec
        Wscript.Echo "Writes Large Per Second: " & objItem.WritesLargePersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

