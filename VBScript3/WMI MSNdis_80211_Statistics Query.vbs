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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSNdis_80211_Statistics", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ACKFailureCount: " & objItem.ACKFailureCount
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "FailedCount: " & objItem.FailedCount
      WScript.Echo "FCSErrorCount: " & objItem.FCSErrorCount
      WScript.Echo "FrameDuplicateCount: " & objItem.FrameDuplicateCount
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MulticastReceivedFrameCount: " & objItem.MulticastReceivedFrameCount
      WScript.Echo "MulticastTransmittedFrameCount: " & objItem.MulticastTransmittedFrameCount
      WScript.Echo "MultipleRetryCount: " & objItem.MultipleRetryCount
      WScript.Echo "ReceivedFragmentCount: " & objItem.ReceivedFragmentCount
      WScript.Echo "RetryCount: " & objItem.RetryCount
      WScript.Echo "RTSFailureCount: " & objItem.RTSFailureCount
      WScript.Echo "RTSSuccessCount: " & objItem.RTSSuccessCount
      WScript.Echo "StatisticsLength: " & objItem.StatisticsLength
      WScript.Echo "TransmittedFragmentCount: " & objItem.TransmittedFragmentCount
      WScript.Echo
   Next
Next

