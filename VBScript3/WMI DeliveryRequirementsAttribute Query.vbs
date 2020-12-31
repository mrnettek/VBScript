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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM DeliveryRequirementsAttribute", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "QueuedDeliveryRequirements: " & objItem.QueuedDeliveryRequirements
      WScript.Echo "RequireOrderedDelivery: " & objItem.RequireOrderedDelivery
      WScript.Echo "TargetContract: " & objItem.TargetContract
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

