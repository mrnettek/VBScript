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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSTapeProblemDeviceError", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "DriveHardwareError: " & objItem.DriveHardwareError
      WScript.Echo "DriveRequiresCleaning: " & objItem.DriveRequiresCleaning
      WScript.Echo "HardError: " & objItem.HardError
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MediaLife: " & objItem.MediaLife
      WScript.Echo "ReadFailure: " & objItem.ReadFailure
      WScript.Echo "ReadWarning: " & objItem.ReadWarning
      WScript.Echo "ScsiInterfaceError: " & objItem.ScsiInterfaceError
      WScript.Echo "TapeSnapped: " & objItem.TapeSnapped
      WScript.Echo "TimetoCleanDrive: " & objItem.TimetoCleanDrive
      WScript.Echo "UnsupportedFormat: " & objItem.UnsupportedFormat
      WScript.Echo "WriteFailure: " & objItem.WriteFailure
      WScript.Echo "WriteWarning: " & objItem.WriteWarning
      WScript.Echo
   Next
Next

