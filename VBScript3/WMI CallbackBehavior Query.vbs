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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM CallbackBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AutomaticSessionShutdown: " & objItem.AutomaticSessionShutdown
      WScript.Echo "ConcurrencyMode: " & objItem.ConcurrencyMode
      WScript.Echo "IgnoreExtensionDataObject: " & objItem.IgnoreExtensionDataObject
      WScript.Echo "IncludeExceptionDetailInFaults: " & objItem.IncludeExceptionDetailInFaults
      WScript.Echo "MaxItemsInObjectGraph: " & objItem.MaxItemsInObjectGraph
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UseSynchronizationContext: " & objItem.UseSynchronizationContext
      WScript.Echo "ValidateMustUnderstand: " & objItem.ValidateMustUnderstand
      WScript.Echo
   Next
Next

