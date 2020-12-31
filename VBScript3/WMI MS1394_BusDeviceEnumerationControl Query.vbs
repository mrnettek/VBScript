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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS1394_BusDeviceEnumerationControl", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      strExclusionElementFlags = Join(objItem.ExclusionElementFlags, ",")
         WScript.Echo "ExclusionElementFlags: " & strExclusionElementFlags
      strExclusionElementList = Join(objItem.ExclusionElementList, ",")
         WScript.Echo "ExclusionElementList: " & strExclusionElementList
      strExclusionElementType = Join(objItem.ExclusionElementType, ",")
         WScript.Echo "ExclusionElementType: " & strExclusionElementType
      WScript.Echo "Flags: " & objItem.Flags
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NumberOfElements: " & objItem.NumberOfElements
      WScript.Echo "Reserved1: " & objItem.Reserved1
      WScript.Echo
   Next
Next

