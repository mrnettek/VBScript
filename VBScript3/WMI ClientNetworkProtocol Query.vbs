On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\Microsoft\SqlServer\ComputerManagement10")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ClientNetworkProtocol", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ProtocolDisplayName: " & objItem.ProtocolDisplayName
      WScript.Echo "ProtocolDLL: " & objItem.ProtocolDLL
      WScript.Echo "ProtocolName: " & objItem.ProtocolName
      WScript.Echo "ProtocolOrder: " & objItem.ProtocolOrder
      WScript.Echo "SupportAlias: " & objItem.SupportAlias
      WScript.Echo
   Next
Next

