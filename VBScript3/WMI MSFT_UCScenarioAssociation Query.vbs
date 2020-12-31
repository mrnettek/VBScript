On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\subscription")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSFT_UCScenarioAssociation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Object: " & objItem.Object
      WScript.Echo "Scenario: " & objItem.Scenario
      WScript.Echo
   Next
Next

