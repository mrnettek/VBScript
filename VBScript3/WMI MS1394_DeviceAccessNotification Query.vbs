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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_DeviceAccessNotification", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BusGeneration: " & objItem.BusGeneration
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "OwnerShipAccessFlags: " & objItem.OwnerShipAccessFlags
      WScript.Echo "RemoteOwnerEUI: " & objItem.RemoteOwnerEUI
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

