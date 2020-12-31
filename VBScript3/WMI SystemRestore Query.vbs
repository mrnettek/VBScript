On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\DEFAULT")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SystemRestore", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "CreationTime: " & objItem.CreationTime
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "EventType: " & objItem.EventType
      WScript.Echo "RestorePointType: " & objItem.RestorePointType
      WScript.Echo "SequenceNumber: " & objItem.SequenceNumber
      WScript.Echo
   Next
Next

