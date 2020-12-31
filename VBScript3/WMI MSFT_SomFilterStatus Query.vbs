On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\Policy")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSFT_SomFilterStatus", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ContainerAvailable: " & objItem.ContainerAvailable
      WScript.Echo "Domain: " & objItem.Domain
      WScript.Echo "SchemaAvailable: " & objItem.SchemaAvailable
      WScript.Echo
   Next
Next

