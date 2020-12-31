' Description: Uses cooked performance counters to monitor the rates of announcements, enumerations, and other browser transmissions


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfNet_Browser").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Announcements Domain Per Second: " & _
            objItem.AnnouncementsDomainPersec
        Wscript.Echo "Announcements Server Per Second: " & _
            objItem.AnnouncementsServerPersec
        Wscript.Echo "Announcements Total Per Second: " & _
            objItem.AnnouncementsTotalPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Duplicate Master Announcements: " & _
            objItem.DuplicateMasterAnnouncements
        Wscript.Echo "Election Packets Per Second: " & _
            objItem.ElectionPacketsPersec
        Wscript.Echo "Enumerations Domain Per Second: " & _
            objItem.EnumerationsDomainPersec
        Wscript.Echo "Enumerations Other Per Second: " & _
            objItem.EnumerationsOtherPersec
        Wscript.Echo "Enumerations Server Per Second: " & _
            objItem.EnumerationsServerPersec
        Wscript.Echo "Enumerations Total Per Second: " & _
            objItem.EnumerationsTotalPersec
        Wscript.Echo "Illegal Datagrams Per Second: " & _
            objItem.IllegalDatagramsPersec
        Wscript.Echo "Mailslot Allocations Failed: " & _
            objItem.MailslotAllocationsFailed
        Wscript.Echo "Mailslot Opens Failed Per Second: " & _
            objItem.MailslotOpensFailedPersec
        Wscript.Echo "Mailslot Receives Failed: " & _
            objItem.MailslotReceivesFailed
        Wscript.Echo "Mailslot Writes Failed: " & objItem.MailslotWritesFailed
        Wscript.Echo "Mailslot Writes Per Second: " & _
            objItem.MailslotWritesPersec
        Wscript.Echo "Missed Mailslot Datagrams: " & _
            objItem.MissedMailslotDatagrams
        Wscript.Echo "Missed Server Announcements: " & _
            objItem.MissedServerAnnouncements
        Wscript.Echo "Missed Server List Requests: " & _
            objItem.MissedServerListRequests
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Server Announce Allocations Failed Per Second: " & _
            objItem.ServerAnnounceAllocationsFailedPersec
        Wscript.Echo "Server List Requests Per Second: " & _
            objItem.ServerListRequestsPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

