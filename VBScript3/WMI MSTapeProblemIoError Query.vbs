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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSTapeProblemIoError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NonMediumErrors: " & objItem.NonMediumErrors
      WScript.Echo "ReadCorrectedWithDelay: " & objItem.ReadCorrectedWithDelay
      WScript.Echo "ReadCorrectedWithoutDelay: " & objItem.ReadCorrectedWithoutDelay
      WScript.Echo "ReadCorrectionAlgorithmProcessed: " & objItem.ReadCorrectionAlgorithmProcessed
      WScript.Echo "ReadTotalCorrectedErrors: " & objItem.ReadTotalCorrectedErrors
      WScript.Echo "ReadTotalErrors: " & objItem.ReadTotalErrors
      WScript.Echo "ReadTotalUncorrectedErrors: " & objItem.ReadTotalUncorrectedErrors
      WScript.Echo "WriteCorrectedWithDelay: " & objItem.WriteCorrectedWithDelay
      WScript.Echo "WriteCorrectedWithoutDelay: " & objItem.WriteCorrectedWithoutDelay
      WScript.Echo "WriteCorrectionAlgorithmProcessed: " & objItem.WriteCorrectionAlgorithmProcessed
      WScript.Echo "WriteTotalCorrectedErrors: " & objItem.WriteTotalCorrectedErrors
      WScript.Echo "WriteTotalErrors: " & objItem.WriteTotalErrors
      WScript.Echo "WriteTotalUncorrectedErrors: " & objItem.WriteTotalUncorrectedErrors
      WScript.Echo
   Next
Next

