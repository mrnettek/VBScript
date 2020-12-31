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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSMouse_PortInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Buttons: " & objItem.Buttons
      WScript.Echo "ConnectorType: " & objItem.ConnectorType
      WScript.Echo "DataQueueSize: " & objItem.DataQueueSize
      WScript.Echo "ErrorCount: " & objItem.ErrorCount
      WScript.Echo "HardwareType: " & objItem.HardwareType
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo
   Next
Next

