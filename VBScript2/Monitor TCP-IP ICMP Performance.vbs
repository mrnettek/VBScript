' Description: Uses cooked performance counters to monitor the rates at which messages are sent and received by using ICMP protocols.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_Tcpip_ICMP").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Messages Outbound Errors: " & _
            objItem.MessagesOutboundErrors
        Wscript.Echo "Messages Per Second: " & objItem.MessagesPersec
        Wscript.Echo "Messages Received Errors: " & _
            objItem.MessagesReceivedErrors
        Wscript.Echo "Messages Received Per Second: " & _
            objItem.MessagesReceivedPersec
        Wscript.Echo "Messages Sent Per Second: " & objItem.MessagesSentPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Received Address Mask: " & objItem.ReceivedAddressMask
        Wscript.Echo "Received Address Mask Reply: " & _
            objItem.ReceivedAddressMaskReply
        Wscript.Echo "Received Destination Unreachable: " & _
            objItem.ReceivedDestUnreachable
        Wscript.Echo "Received Echo Per Second: " & objItem.ReceivedEchoPersec
        Wscript.Echo "Received Echo Reply Per Second: " & _
            objItem.ReceivedEchoReplyPersec
        Wscript.Echo "Received Parameter Problem: " & _
            objItem.ReceivedParameterProblem
        Wscript.Echo "Received Redirect Per Second: " & _
            objItem.ReceivedRedirectPersec
        Wscript.Echo "Received Source Quench: " & _
            objItem.ReceivedSourceQuench
        Wscript.Echo "Received Time Exceeded: " & _
        objItem.ReceivedTimeExceeded
        Wscript.Echo "Received Timestamp Per Second: " & _
            objItem.ReceivedTimestampPersec
        Wscript.Echo "Received Timestamp Reply Per Second: " & _
            objItem.ReceivedTimestampReplyPersec
        Wscript.Echo "Sent Address Mask: " & objItem.SentAddressMask
        Wscript.Echo "Sent Address Mask Reply: " & _
            objItem.SentAddressMaskReply
        Wscript.Echo "Sent Destination Unreachable: " & _
            objItem.SentDestinationUnreachable
        Wscript.Echo "Sent Echo Per Second: " & objItem.SentEchoPersec
        Wscript.Echo "Sent Echo Reply Per Second: " & _
            objItem.SentEchoReplyPersec
        Wscript.Echo "Sent Parameter Problem: " & objItem.SentParameterProblem
        Wscript.Echo "Sent Redirect Per Second: " & objItem.SentRedirectPersec
        Wscript.Echo "Sent Source Quench: " & objItem.SentSourceQuench
        Wscript.Echo "Sent Time Exceeded: " & objItem.SentTimeExceeded
        Wscript.Echo "Sent Timestamp Per Second: " & _
            objItem.SentTimestampPersec
        Wscript.Echo "Sent Timestamp Reply Per Second: " & _
            objItem.SentTimestampReplyPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

