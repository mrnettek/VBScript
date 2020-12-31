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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM AppDomainInfo", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AppDomainId: " & objItem.AppDomainId
      WScript.Echo "IsDefault: " & objItem.IsDefault
      WScript.Echo "LogMalformedMessages: " & objItem.LogMalformedMessages
      WScript.Echo "LogMessagesAtServiceLevel: " & objItem.LogMessagesAtServiceLevel
      WScript.Echo "LogMessagesAtTransportLevel: " & objItem.LogMessagesAtTransportLevel
      strMessageLoggingTraceListeners = Join(objItem.MessageLoggingTraceListeners, ",")
         WScript.Echo "MessageLoggingTraceListeners: " & strMessageLoggingTraceListeners
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "PerformanceCounters: " & objItem.PerformanceCounters
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo "ServiceConfigPath: " & objItem.ServiceConfigPath
      strServiceModelTraceListeners = Join(objItem.ServiceModelTraceListeners, ",")
         WScript.Echo "ServiceModelTraceListeners: " & strServiceModelTraceListeners
      WScript.Echo "TraceLevel: " & objItem.TraceLevel
      WScript.Echo
   Next
Next

