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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Endpoint", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Address: " & objItem.Address
      strAddressHeaders = Join(objItem.AddressHeaders, ",")
         WScript.Echo "AddressHeaders: " & strAddressHeaders
      WScript.Echo "AddressIdentity: " & objItem.AddressIdentity
      WScript.Echo "AppDomainId: " & objItem.AppDomainId
      strBehaviors = Join(objItem.Behaviors, ",")
         WScript.Echo "Behaviors: " & strBehaviors
      WScript.Echo "Binding: " & objItem.Binding
      WScript.Echo "Contract: " & objItem.Contract
      WScript.Echo "ContractName: " & objItem.ContractName
      WScript.Echo "CounterInstanceName: " & objItem.CounterInstanceName
      WScript.Echo "ListenUri: " & objItem.ListenUri
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo
   Next
Next

