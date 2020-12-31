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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM BatteryStatus", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "ChargeRate: " & objItem.ChargeRate
      WScript.Echo "Charging: " & objItem.Charging
      WScript.Echo "Critical: " & objItem.Critical
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DischargeRate: " & objItem.DischargeRate
      WScript.Echo "Discharging: " & objItem.Discharging
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "PowerOnline: " & objItem.PowerOnline
      WScript.Echo "RemainingCapacity: " & objItem.RemainingCapacity
      WScript.Echo "Tag: " & objItem.Tag
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "Voltage: " & objItem.Voltage
      WScript.Echo
   Next
Next

