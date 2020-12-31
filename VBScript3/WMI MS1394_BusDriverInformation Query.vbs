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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_BusDriverInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BusDDIVersion: " & objItem.BusDDIVersion
      strConfigRom = Join(objItem.ConfigRom, ",")
         WScript.Echo "ConfigRom: " & strConfigRom
      WScript.Echo "ConfigRomSize: " & objItem.ConfigRomSize
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "LocalHostControllerEUI: " & objItem.LocalHostControllerEUI
      WScript.Echo "MaxPhySpeed: " & objItem.MaxPhySpeed
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo
   Next
Next

