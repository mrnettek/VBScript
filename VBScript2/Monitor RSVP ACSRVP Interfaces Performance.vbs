' Description: Uses cooked performance counters to monitor RSVP or ACS service performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_RSVP_RSVPService").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Bytes in QoS Notifications: " & _
            objItem.BytesinQoSnotifications
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Failed QoS Requests: " & objItem.FailedQoSrequests
        Wscript.Echo "Failed QoS Sends: " & objItem.FailedQoSsends
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Network Interfaces: " & objItem.NetworkInterfaces
        Wscript.Echo "Network Sockets: " & objItem.Networksockets
        Wscript.Echo "QoS Clients: " & objItem.QoSclients
        Wscript.Echo "QoS-Enabled Receivers: " & objItem.QoSenabledreceivers
        Wscript.Echo "QoS-Enabled Senders: " & objItem.QoSenabledsenders
        Wscript.Echo "QoS Notifications: " & objItem.QoSnotifications
        Wscript.Echo "RSVP Sessions: " & objItem.RSVPsessions
        Wscript.Echo "Timers: " & objItem.Timers
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

