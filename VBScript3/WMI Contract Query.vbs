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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Contract", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AppDomainId: " & objItem.AppDomainId
      strBehaviors = Join(objItem.Behaviors, ",")
         WScript.Echo "Behaviors: " & strBehaviors
      WScript.Echo "CallbackContract: " & objItem.CallbackContract
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Namespace: " & objItem.Namespace
      strOperations = Join(objItem.Operations, ",")
         WScript.Echo "Operations: " & strOperations
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo "SessionMode: " & objItem.SessionMode
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

