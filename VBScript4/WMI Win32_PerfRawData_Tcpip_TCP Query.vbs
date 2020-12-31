On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_TCP",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConnectionFailures: " & objItem.ConnectionFailures
    Wscript.Echo "ConnectionsActive: " & objItem.ConnectionsActive
    Wscript.Echo "ConnectionsEstablished: " & objItem.ConnectionsEstablished
    Wscript.Echo "ConnectionsPassive: " & objItem.ConnectionsPassive
    Wscript.Echo "ConnectionsReset: " & objItem.ConnectionsReset
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SegmentsPersec: " & objItem.SegmentsPersec
    Wscript.Echo "SegmentsReceivedPersec: " & objItem.SegmentsReceivedPersec
    Wscript.Echo "SegmentsRetransmittedPersec: " & objItem.SegmentsRetransmittedPersec
    Wscript.Echo "SegmentsSentPersec: " & objItem.SegmentsSentPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

