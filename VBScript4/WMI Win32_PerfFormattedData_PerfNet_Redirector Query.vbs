On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfNet_Redirector",,48)
For Each objItem in colItems
    Wscript.Echo "BytesReceivedPersec: " & objItem.BytesReceivedPersec
    Wscript.Echo "BytesTotalPersec: " & objItem.BytesTotalPersec
    Wscript.Echo "BytesTransmittedPersec: " & objItem.BytesTransmittedPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConnectsCore: " & objItem.ConnectsCore
    Wscript.Echo "ConnectsLanManager20: " & objItem.ConnectsLanManager20
    Wscript.Echo "ConnectsLanManager21: " & objItem.ConnectsLanManager21
    Wscript.Echo "ConnectsWindowsNT: " & objItem.ConnectsWindowsNT
    Wscript.Echo "CurrentCommands: " & objItem.CurrentCommands
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FileDataOperationsPersec: " & objItem.FileDataOperationsPersec
    Wscript.Echo "FileReadOperationsPersec: " & objItem.FileReadOperationsPersec
    Wscript.Echo "FileWriteOperationsPersec: " & objItem.FileWriteOperationsPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NetworkErrorsPersec: " & objItem.NetworkErrorsPersec
    Wscript.Echo "PacketsPersec: " & objItem.PacketsPersec
    Wscript.Echo "PacketsReceivedPersec: " & objItem.PacketsReceivedPersec
    Wscript.Echo "PacketsTransmittedPersec: " & objItem.PacketsTransmittedPersec
    Wscript.Echo "ReadBytesCachePersec: " & objItem.ReadBytesCachePersec
    Wscript.Echo "ReadBytesNetworkPersec: " & objItem.ReadBytesNetworkPersec
    Wscript.Echo "ReadBytesNonPagingPersec: " & objItem.ReadBytesNonPagingPersec
    Wscript.Echo "ReadBytesPagingPersec: " & objItem.ReadBytesPagingPersec
    Wscript.Echo "ReadOperationsRandomPersec: " & objItem.ReadOperationsRandomPersec
    Wscript.Echo "ReadPacketsPersec: " & objItem.ReadPacketsPersec
    Wscript.Echo "ReadPacketsSmallPersec: " & objItem.ReadPacketsSmallPersec
    Wscript.Echo "ReadsDeniedPersec: " & objItem.ReadsDeniedPersec
    Wscript.Echo "ReadsLargePersec: " & objItem.ReadsLargePersec
    Wscript.Echo "ServerDisconnects: " & objItem.ServerDisconnects
    Wscript.Echo "ServerReconnects: " & objItem.ServerReconnects
    Wscript.Echo "ServerSessions: " & objItem.ServerSessions
    Wscript.Echo "ServerSessionsHung: " & objItem.ServerSessionsHung
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "WriteBytesCachePersec: " & objItem.WriteBytesCachePersec
    Wscript.Echo "WriteBytesNetworkPersec: " & objItem.WriteBytesNetworkPersec
    Wscript.Echo "WriteBytesNonPagingPersec: " & objItem.WriteBytesNonPagingPersec
    Wscript.Echo "WriteBytesPagingPersec: " & objItem.WriteBytesPagingPersec
    Wscript.Echo "WriteOperationsRandomPersec: " & objItem.WriteOperationsRandomPersec
    Wscript.Echo "WritePacketsPersec: " & objItem.WritePacketsPersec
    Wscript.Echo "WritePacketsSmallPersec: " & objItem.WritePacketsSmallPersec
    Wscript.Echo "WritesDeniedPersec: " & objItem.WritesDeniedPersec
    Wscript.Echo "WritesLargePersec: " & objItem.WritesLargePersec
Next

