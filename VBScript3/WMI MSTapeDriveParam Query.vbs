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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSTapeDriveParam", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "CompressionCapable: " & objItem.CompressionCapable
      WScript.Echo "CompressionEnabled: " & objItem.CompressionEnabled
      WScript.Echo "DefaultBlockSize: " & objItem.DefaultBlockSize
      WScript.Echo "HardwareErrorCorrection: " & objItem.HardwareErrorCorrection
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MaximumBlockSize: " & objItem.MaximumBlockSize
      WScript.Echo "MaximumPartitionCount: " & objItem.MaximumPartitionCount
      WScript.Echo "MinimumBlockSize: " & objItem.MinimumBlockSize
      WScript.Echo "ReportSetmarks: " & objItem.ReportSetmarks
      WScript.Echo
   Next
Next

