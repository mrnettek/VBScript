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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SqlService", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AcceptPause: " & objItem.AcceptPause
      WScript.Echo "AcceptStop: " & objItem.AcceptStop
      WScript.Echo "BinaryPath: " & objItem.BinaryPath
      strDependencies = Join(objItem.Dependencies, ",")
         WScript.Echo "Dependencies: " & strDependencies
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DisplayName: " & objItem.DisplayName
      WScript.Echo "ErrorControl: " & objItem.ErrorControl
      WScript.Echo "ExitCode: " & objItem.ExitCode
      WScript.Echo "HostName: " & objItem.HostName
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo "ServiceName: " & objItem.ServiceName
      WScript.Echo "SQLServiceType: " & objItem.SQLServiceType
      WScript.Echo "StartMode: " & objItem.StartMode
      WScript.Echo "StartName: " & objItem.StartName
      WScript.Echo "State: " & objItem.State
      WScript.Echo
   Next
Next

