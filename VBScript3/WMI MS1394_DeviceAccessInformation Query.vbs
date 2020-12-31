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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_DeviceAccessInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NotificationFlags: " & objItem.NotificationFlags
      WScript.Echo "OwnershipAccessFlags: " & objItem.OwnershipAccessFlags
      WScript.Echo "RemoteOwnerEUI: " & objItem.RemoteOwnerEUI
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

