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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM WorkflowServiceBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AddressFilterMode: " & objItem.AddressFilterMode
      WScript.Echo "ConfigurationName: " & objItem.ConfigurationName
      WScript.Echo "IgnoreExtensionDataObject: " & objItem.IgnoreExtensionDataObject
      WScript.Echo "IncludeExceptionDetailInFaults: " & objItem.IncludeExceptionDetailInFaults
      WScript.Echo "MaxItemsInObjectGraph: " & objItem.MaxItemsInObjectGraph
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Namespace: " & objItem.Namespace
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UseSynchronizationContext: " & objItem.UseSynchronizationContext
      WScript.Echo "ValidateMustUnderstand: " & objItem.ValidateMustUnderstand
      WScript.Echo "WorkflowDefinitionPath: " & objItem.WorkflowDefinitionPath
      WScript.Echo "WorkflowRulesPath: " & objItem.WorkflowRulesPath
      WScript.Echo "WorkflowType: " & objItem.WorkflowType
      WScript.Echo
   Next
Next

