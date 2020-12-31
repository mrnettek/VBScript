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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSAcpi_ThermalZoneTemperature", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      strActiveTripPoint = Join(objItem.ActiveTripPoint, ",")
         WScript.Echo "ActiveTripPoint: " & strActiveTripPoint
      WScript.Echo "ActiveTripPointCount: " & objItem.ActiveTripPointCount
      WScript.Echo "CriticalTripPoint: " & objItem.CriticalTripPoint
      WScript.Echo "CurrentTemperature: " & objItem.CurrentTemperature
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "PassiveTripPoint: " & objItem.PassiveTripPoint
      WScript.Echo "Reserved: " & objItem.Reserved
      WScript.Echo "SamplingPeriod: " & objItem.SamplingPeriod
      WScript.Echo "ThermalConstant1: " & objItem.ThermalConstant1
      WScript.Echo "ThermalConstant2: " & objItem.ThermalConstant2
      WScript.Echo "ThermalStamp: " & objItem.ThermalStamp
      WScript.Echo
   Next
Next

