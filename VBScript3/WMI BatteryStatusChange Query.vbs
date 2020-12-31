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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM BatteryStatusChange", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Charging: " & objItem.Charging
      WScript.Echo "Critical: " & objItem.Critical
      WScript.Echo "Discharging: " & objItem.Discharging
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "PowerOnline: " & objItem.PowerOnline
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "Tag: " & objItem.Tag
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo
   Next
Next

