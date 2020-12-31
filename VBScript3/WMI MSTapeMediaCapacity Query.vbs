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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSTapeMediaCapacity", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "AvailableCapacity: " & objItem.AvailableCapacity
      WScript.Echo "BlockSize: " & objItem.BlockSize
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MaximumCapacity: " & objItem.MaximumCapacity
      WScript.Echo "MediaWriteProtected: " & objItem.MediaWriteProtected
      WScript.Echo "PartitionCount: " & objItem.PartitionCount
      WScript.Echo
   Next
Next

