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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MsmqBindingElementBase", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "CustomDeadLetterQueue: " & objItem.CustomDeadLetterQueue
      WScript.Echo "DeadLetterQueue: " & objItem.DeadLetterQueue
      WScript.Echo "Durable: " & objItem.Durable
      WScript.Echo "ExactlyOnce: " & objItem.ExactlyOnce
      WScript.Echo "ManualAddressing: " & objItem.ManualAddressing
      WScript.Echo "MaxBufferPoolSize: " & objItem.MaxBufferPoolSize
      WScript.Echo "MaxReceivedMessageSize: " & objItem.MaxReceivedMessageSize
      WScript.Echo "MaxRetryCycles: " & objItem.MaxRetryCycles
      WScript.Echo "ReceiveErrorHandling: " & objItem.ReceiveErrorHandling
      WScript.Echo "ReceiveRetryCount: " & objItem.ReceiveRetryCount
      WScript.Echo "RetryCycleDelay: " & WMIDateStringToDate(objItem.RetryCycleDelay)
      WScript.Echo "Scheme: " & objItem.Scheme
      WScript.Echo "TimeToLive: " & WMIDateStringToDate(objItem.TimeToLive)
      WScript.Echo "UseMsmqTracing: " & objItem.UseMsmqTracing
      WScript.Echo "UseSourceJournal: " & objItem.UseSourceJournal
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

