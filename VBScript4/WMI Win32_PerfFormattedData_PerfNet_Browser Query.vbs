On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfNet_Browser",,48)
For Each objItem in colItems
    Wscript.Echo "AnnouncementsDomainPersec: " & objItem.AnnouncementsDomainPersec
    Wscript.Echo "AnnouncementsServerPersec: " & objItem.AnnouncementsServerPersec
    Wscript.Echo "AnnouncementsTotalPersec: " & objItem.AnnouncementsTotalPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DuplicateMasterAnnouncements: " & objItem.DuplicateMasterAnnouncements
    Wscript.Echo "ElectionPacketsPersec: " & objItem.ElectionPacketsPersec
    Wscript.Echo "EnumerationsDomainPersec: " & objItem.EnumerationsDomainPersec
    Wscript.Echo "EnumerationsOtherPersec: " & objItem.EnumerationsOtherPersec
    Wscript.Echo "EnumerationsServerPersec: " & objItem.EnumerationsServerPersec
    Wscript.Echo "EnumerationsTotalPersec: " & objItem.EnumerationsTotalPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "IllegalDatagramsPersec: " & objItem.IllegalDatagramsPersec
    Wscript.Echo "MailslotAllocationsFailed: " & objItem.MailslotAllocationsFailed
    Wscript.Echo "MailslotOpensFailedPersec: " & objItem.MailslotOpensFailedPersec
    Wscript.Echo "MailslotReceivesFailed: " & objItem.MailslotReceivesFailed
    Wscript.Echo "MailslotWritesFailed: " & objItem.MailslotWritesFailed
    Wscript.Echo "MailslotWritesPersec: " & objItem.MailslotWritesPersec
    Wscript.Echo "MissedMailslotDatagrams: " & objItem.MissedMailslotDatagrams
    Wscript.Echo "MissedServerAnnouncements: " & objItem.MissedServerAnnouncements
    Wscript.Echo "MissedServerListRequests: " & objItem.MissedServerListRequests
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ServerAnnounceAllocationsFailedPersec: " & objItem.ServerAnnounceAllocationsFailedPersec
    Wscript.Echo "ServerListRequestsPersec: " & objItem.ServerListRequestsPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

