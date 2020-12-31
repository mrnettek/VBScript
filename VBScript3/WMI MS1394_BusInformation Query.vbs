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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_BusInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Generation: " & objItem.Generation
      WScript.Echo "InstanceName: " & objItem.InstanceName
      strLocalHostSelfId = Join(objItem.LocalHostSelfId, ",")
         WScript.Echo "LocalHostSelfId: " & strLocalHostSelfId
      WScript.Echo "Reserved1: " & objItem.Reserved1
      strTopologyMap = Join(objItem.TopologyMap, ",")
         WScript.Echo "TopologyMap: " & strTopologyMap
      strTreeTopologyMap = Join(objItem.TreeTopologyMap, ",")
         WScript.Echo "TreeTopologyMap: " & strTreeTopologyMap
      WScript.Echo
   Next
Next

