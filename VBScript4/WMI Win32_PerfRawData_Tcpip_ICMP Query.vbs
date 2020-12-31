On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_ICMP",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "MessagesOutboundErrors: " & objItem.MessagesOutboundErrors
    Wscript.Echo "MessagesPersec: " & objItem.MessagesPersec
    Wscript.Echo "MessagesReceivedErrors: " & objItem.MessagesReceivedErrors
    Wscript.Echo "MessagesReceivedPersec: " & objItem.MessagesReceivedPersec
    Wscript.Echo "MessagesSentPersec: " & objItem.MessagesSentPersec
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ReceivedAddressMask: " & objItem.ReceivedAddressMask
    Wscript.Echo "ReceivedAddressMaskReply: " & objItem.ReceivedAddressMaskReply
    Wscript.Echo "ReceivedDestUnreachable: " & objItem.ReceivedDestUnreachable
    Wscript.Echo "ReceivedEchoPersec: " & objItem.ReceivedEchoPersec
    Wscript.Echo "ReceivedEchoReplyPersec: " & objItem.ReceivedEchoReplyPersec
    Wscript.Echo "ReceivedParameterProblem: " & objItem.ReceivedParameterProblem
    Wscript.Echo "ReceivedRedirectPersec: " & objItem.ReceivedRedirectPersec
    Wscript.Echo "ReceivedSourceQuench: " & objItem.ReceivedSourceQuench
    Wscript.Echo "ReceivedTimeExceeded: " & objItem.ReceivedTimeExceeded
    Wscript.Echo "ReceivedTimestampPersec: " & objItem.ReceivedTimestampPersec
    Wscript.Echo "ReceivedTimestampReplyPersec: " & objItem.ReceivedTimestampReplyPersec
    Wscript.Echo "SentAddressMask: " & objItem.SentAddressMask
    Wscript.Echo "SentAddressMaskReply: " & objItem.SentAddressMaskReply
    Wscript.Echo "SentDestinationUnreachable: " & objItem.SentDestinationUnreachable
    Wscript.Echo "SentEchoPersec: " & objItem.SentEchoPersec
    Wscript.Echo "SentEchoReplyPersec: " & objItem.SentEchoReplyPersec
    Wscript.Echo "SentParameterProblem: " & objItem.SentParameterProblem
    Wscript.Echo "SentRedirectPersec: " & objItem.SentRedirectPersec
    Wscript.Echo "SentSourceQuench: " & objItem.SentSourceQuench
    Wscript.Echo "SentTimeExceeded: " & objItem.SentTimeExceeded
    Wscript.Echo "SentTimestampPersec: " & objItem.SentTimestampPersec
    Wscript.Echo "SentTimestampReplyPersec: " & objItem.SentTimestampReplyPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

