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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM FilestreamSettings", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccessLevel: " & objItem.AccessLevel
      WScript.Echo "IncompleteOperation: " & objItem.IncompleteOperation
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "IsClustered: " & objItem.IsClustered
      WScript.Echo "RsFxVersion: " & objItem.RsFxVersion
      WScript.Echo "ShareName: " & objItem.ShareName
      WScript.Echo
   Next
Next

