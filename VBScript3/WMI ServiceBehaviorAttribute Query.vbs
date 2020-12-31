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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceBehaviorAttribute", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AddressFilterMode: " & objItem.AddressFilterMode
      WScript.Echo "AutomaticSessionShutdown: " & objItem.AutomaticSessionShutdown
      WScript.Echo "ConcurrencyMode: " & objItem.ConcurrencyMode
      WScript.Echo "ConfigurationName: " & objItem.ConfigurationName
      WScript.Echo "IgnoreExtensionDataObject: " & objItem.IgnoreExtensionDataObject
      WScript.Echo "IncludeExceptionDetailInFaults: " & objItem.IncludeExceptionDetailInFaults
      WScript.Echo "InstanceContextMode: " & objItem.InstanceContextMode
      WScript.Echo "MaxItemsInObjectGraph: " & objItem.MaxItemsInObjectGraph
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Namespace: " & objItem.Namespace
      WScript.Echo "ReleaseServiceInstanceOnTransactionComplete: " & objItem.ReleaseServiceInstanceOnTransactionComplete
      WScript.Echo "TransactionAutoCompleteOnSessionClose: " & objItem.TransactionAutoCompleteOnSessionClose
      WScript.Echo "TransactionIsolationLevel: " & objItem.TransactionIsolationLevel
      WScript.Echo "TransactionTimeout: " & WMIDateStringToDate(objItem.TransactionTimeout)
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UseSynchronizationContext: " & objItem.UseSynchronizationContext
      WScript.Echo "ValidateMustUnderstand: " & objItem.ValidateMustUnderstand
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

