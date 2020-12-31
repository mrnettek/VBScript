On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_NetworkInterface",,48)
For Each objItem in colItems
    Wscript.Echo "BytesReceivedPersec: " & objItem.BytesReceivedPersec
    Wscript.Echo "BytesSentPersec: " & objItem.BytesSentPersec
    Wscript.Echo "BytesTotalPersec: " & objItem.BytesTotalPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CurrentBandwidth: " & objItem.CurrentBandwidth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OutputQueueLength: " & objItem.OutputQueueLength
    Wscript.Echo "PacketsOutboundDiscarded: " & objItem.PacketsOutboundDiscarded
    Wscript.Echo "PacketsOutboundErrors: " & objItem.PacketsOutboundErrors
    Wscript.Echo "PacketsPersec: " & objItem.PacketsPersec
    Wscript.Echo "PacketsReceivedDiscarded: " & objItem.PacketsReceivedDiscarded
    Wscript.Echo "PacketsReceivedErrors: " & objItem.PacketsReceivedErrors
    Wscript.Echo "PacketsReceivedNonUnicastPersec: " & objItem.PacketsReceivedNonUnicastPersec
    Wscript.Echo "PacketsReceivedPersec: " & objItem.PacketsReceivedPersec
    Wscript.Echo "PacketsReceivedUnicastPersec: " & objItem.PacketsReceivedUnicastPersec
    Wscript.Echo "PacketsReceivedUnknown: " & objItem.PacketsReceivedUnknown
    Wscript.Echo "PacketsSentNonUnicastPersec: " & objItem.PacketsSentNonUnicastPersec
    Wscript.Echo "PacketsSentPersec: " & objItem.PacketsSentPersec
    Wscript.Echo "PacketsSentUnicastPersec: " & objItem.PacketsSentUnicastPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

