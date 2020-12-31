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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_BusErrorInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Generation: " & objItem.Generation
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strNonEnumeratedDevices = Join(objItem.NonEnumeratedDevices, ",")
         WScript.Echo "NonEnumeratedDevices: " & strNonEnumeratedDevices
      WScript.Echo "NumberOfNonEnumeratedDevices: " & objItem.NumberOfNonEnumeratedDevices
      WScript.Echo "NumberOfUnpoweredDevices: " & objItem.NumberOfUnpoweredDevices
      WScript.Echo "Reserved1: " & objItem.Reserved1
      strUnpoweredDevices = Join(objItem.UnpoweredDevices, ",")
         WScript.Echo "UnpoweredDevices: " & strUnpoweredDevices
      WScript.Echo
   Next
Next

