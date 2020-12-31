' Description: Uses cooked performance counters to monitor the number of local network interfaces visible to, and used by the RSVP service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_RSVP_RVPInterfaces").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Blocked RESVs: " & objItem.BlockedRESVs
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "General Failures: " & objItem.Generalfailures
        Wscript.Echo "Maximum Admitted Bandwidth: " & _
            objItem.Maximumadmittedbandwidth
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Number of Active Flows: " & objItem.Numberofactiveflows
        Wscript.Echo "Number of Incoming Messages Dropped: " & _ 
            objItem.Numberofincomingmessagesdropped
        Wscript.Echo "Number of Outgoing Messages Dropped: " & _
            objItem.Numberofoutgoingmessagesdropped
        Wscript.Echo "PATH Error Messages Received: " & _
            objItem.PATHERRmessagesreceived
        Wscript.Echo "PATH Error Messages Sent: " & _
            objItem.PATHERRmessagessent
        Wscript.Echo "PATH Messages Received: " & objItem.PATHmessagesreceived
        Wscript.Echo "PATH Messages Sent: " & objItem.PATHmessagessent
        Wscript.Echo "PATH State Block Timeouts: " & _
            objItem.PATHstateblocktimeouts
        Wscript.Echo "PATH TEAR Messages Received: " & _
            objItem.PATHTEARmessagesreceived
        Wscript.Echo "PATH TEAR Messages Sent: " & _
            objItem.PATHTEARmessagessent
        Wscript.Echo "Policy Control Failures: " & _
        objItem.Policycontrolfailures
        Wscript.Echo "Receive Messages Errors Big Messages: " & _
            objItem.ReceivemessageserrorsBigmessages
        Wscript.Echo "Receive Messages Errors No Memory: " & _
             objItem.ReceivemessageserrorsNomemory
        Wscript.Echo "Reserved Bandwidth: " & objItem.Reservedbandwidth
        Wscript.Echo "Resource Control Failures: " & _
            objItem.Resourcecontrolfailures
        Wscript.Echo "RESV CONFIRM Messages Received: " & _
            objItem.RESVCONFIRMmessagesreceived
        Wscript.Echo "RESV CONFIRM Messages Sent: " & _
            objItem.RESVCONFIRMmessagessent
        Wscript.Echo "RESV Error Messages Received: " & _
            objItem.RESVERRmessagesreceived
        Wscript.Echo "RESV Error Messages Sent: " & _
            objItem.RESVERRmessagessent
        Wscript.Echo "RESV Messages Received: " & objItem.RESVmessagesreceived
        Wscript.Echo "RESV Messages Sent: " & objItem.RESVmessagessent
        Wscript.Echo "RESV State Block Timeouts: " & _
            objItem.RESVstateblocktimeouts
        Wscript.Echo "RESV TEAR Messages Received: " & _
            objItem.RESVTEARmessagesreceived
        Wscript.Echo "RESV TEAR Messages Sent: " & _
            objItem.RESVTEARmessagessent
        Wscript.Echo "Send Messages Errors Big Messages: " & _
            objItem.SendmessageserrorsBigmessages
        Wscript.Echo "Send Messages Errors No Memory: " & _
            objItem.SendmessageserrorsNomemory
        Wscript.Echo "Signaling Bytes Received: " & _
            objItem.Signalingbytesreceived
        Wscript.Echo "Signaling Bytes Sent: " & objItem.Signalingbytessent
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

