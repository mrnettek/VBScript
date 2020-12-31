' Description: Uses cooked performance counters to monitor the rates at which IP datagrams are sent and received by using IP protocols.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCPIP_IP").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Datagrams Forwarded Per Second: " & _
            objItem.DatagramsForwardedPersec
        Wscript.Echo "Datagrams Outbound Discarded: " & _
            objItem.DatagramsOutboundDiscarded
        Wscript.Echo "Datagrams Outbound No Route: " & _
            objItem.DatagramsOutboundNoRoute
        Wscript.Echo "Datagrams Per Second: " & objItem.DatagramsPersec
        Wscript.Echo "Datagrams Received Address Errors: " & _
            objItem.DatagramsReceivedAddressErrors
        Wscript.Echo "Datagrams Received Delivered Per Second: " & _
            objItem.DatagramsReceivedDeliveredPersec
        Wscript.Echo "Datagrams Received Discarded: " & _
            objItem.DatagramsReceivedDiscarded
        Wscript.Echo "Datagrams Received Header Errors: " & _
            objItem.DatagramsReceivedHeaderErrors
        Wscript.Echo "Datagrams Received Per Second: " & _
            objItem.DatagramsReceivedPersec
        Wscript.Echo "Datagrams Received Unknown Protocol: " & _
            objItem.DatagramsReceivedUnknownProtocol
        Wscript.Echo "Datagrams Sent Per Second: " & _
            objItem.DatagramsSentPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Fragmentation Failures: " & _
            objItem.FragmentationFailures
        Wscript.Echo "Fragmented Datagrams Per Second: " & _
            objItem.FragmentedDatagramsPersec
        Wscript.Echo "Fragment Reassembly Failures: " & _
            objItem.FragmentReassemblyFailures
        Wscript.Echo "Fragments Created Per Second: " & _
            objItem.FragmentsCreatedPersec
        Wscript.Echo "Fragments Reassembled Per Second: " & _
            objItem.FragmentsReassembledPersec
        Wscript.Echo "Fragments Received Per Second: " & _
            objItem.FragmentsReceivedPersec
        Wscript.Echo "Name: " & objItem.Name
        objRefresher.Refresh
    Next
Next

