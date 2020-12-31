On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ReliableSessionBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AcknowledgementInterval: " & WMIDateStringToDate(objItem.AcknowledgementInterval)
      WScript.Echo "FlowControlEnabled: " & objItem.FlowControlEnabled
      WScript.Echo "InactivityTimeout: " & WMIDateStringToDate(objItem.InactivityTimeout)
      WScript.Echo "MaxPendingChannels: " & objItem.MaxPendingChannels
      WScript.Echo "MaxRetryCount: " & objItem.MaxRetryCount
      WScript.Echo "MaxTransferWindowSize: " & objItem.MaxTransferWindowSize
      WScript.Echo "Ordered: " & objItem.Ordered
      WScript.Echo "ReliableMessagingVersion: " & objItem.ReliableMessagingVersion
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

