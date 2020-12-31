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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM OperationBehaviorAttribute", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AutoDisposeParameters: " & objItem.AutoDisposeParameters
      WScript.Echo "Impersonation: " & objItem.Impersonation
      WScript.Echo "ReleaseInstanceMode: " & objItem.ReleaseInstanceMode
      WScript.Echo "TransactionAutoComplete: " & objItem.TransactionAutoComplete
      WScript.Echo "TransactionScopeRequired: " & objItem.TransactionScopeRequired
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

