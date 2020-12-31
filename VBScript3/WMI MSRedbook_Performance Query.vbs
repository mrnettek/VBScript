On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSRedbook_Performance", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "DataProcessed: " & objItem.DataProcessed
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "StreamPausedCount: " & objItem.StreamPausedCount
      WScript.Echo "TimeReadDelay: " & objItem.TimeReadDelay
      WScript.Echo "TimeReading: " & objItem.TimeReading
      WScript.Echo "TimeStreamDelay: " & objItem.TimeStreamDelay
      WScript.Echo "TimeStreaming: " & objItem.TimeStreaming
      WScript.Echo
   Next
Next

