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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSerial_PerformanceInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BufferOverrunErrorCount: " & objItem.BufferOverrunErrorCount
      WScript.Echo "FrameErrorCount: " & objItem.FrameErrorCount
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "ParityErrorCount: " & objItem.ParityErrorCount
      WScript.Echo "ReceivedCount: " & objItem.ReceivedCount
      WScript.Echo "SerialOverrunErrorCount: " & objItem.SerialOverrunErrorCount
      WScript.Echo "TransmittedCount: " & objItem.TransmittedCount
      WScript.Echo
   Next
Next

