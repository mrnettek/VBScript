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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_DeviceInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      strConfigRomHeader = Join(objItem.ConfigRomHeader, ",")
         WScript.Echo "ConfigRomHeader: " & strConfigRomHeader
      WScript.Echo "DeviceEUI: " & objItem.DeviceEUI
      WScript.Echo "DeviceType: " & objItem.DeviceType
      WScript.Echo "Generation: " & objItem.Generation
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NodeAddress: " & objItem.NodeAddress
      WScript.Echo "PhyDelay: " & objItem.PhyDelay
      WScript.Echo "PhySpeed: " & objItem.PhySpeed
      WScript.Echo "PowerClass: " & objItem.PowerClass
      WScript.Echo "Reserved1: " & objItem.Reserved1
      strSelfId = Join(objItem.SelfId, ",")
         WScript.Echo "SelfId: " & strSelfId
      strUnitDirectory = Join(objItem.UnitDirectory, ",")
         WScript.Echo "UnitDirectory: " & strUnitDirectory
      WScript.Echo
   Next
Next

