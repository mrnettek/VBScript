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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM DataContractSerializerOperationBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "IgnoreExtensionDataObject: " & objItem.IgnoreExtensionDataObject
      WScript.Echo "MaxItemsInObjectGraph: " & objItem.MaxItemsInObjectGraph
      WScript.Echo "Style: " & objItem.Style
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

