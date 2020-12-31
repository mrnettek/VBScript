' Description: Uses cooked performance counters to monitor the telephony service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TAPISrv_Telephony").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Active Lines: " & objItem.ActiveLines
        Wscript.Echo "Active Telephones: " & objItem.ActiveTelephones
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Client Applications: " & objItem.ClientApps
        Wscript.Echo "Current Incoming Calls: " & objItem.CurrentIncomingCalls
        Wscript.Echo "Current Outgoing Calls: " & objItem.CurrentOutgoingCalls
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Incoming Calls Per Second: " & _
            objItem.IncomingCallsPersec
        Wscript.Echo "Lines: " & objItem.Lines
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Outgoing Calls Per Second: " & _
        objItem.OutgoingCallsPersec
        Wscript.Echo "Telephone Devices: " & objItem.TelephoneDevices
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

