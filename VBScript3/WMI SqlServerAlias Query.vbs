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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SqlServerAlias", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AliasName: " & objItem.AliasName
      WScript.Echo "ConnectionString: " & objItem.ConnectionString
      WScript.Echo "ProtocolName: " & objItem.ProtocolName
      WScript.Echo "ServerName: " & objItem.ServerName
      WScript.Echo
   Next
Next

