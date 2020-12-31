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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM BatteryStaticData", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Capabilities: " & objItem.Capabilities
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Chemistry: " & objItem.Chemistry
      WScript.Echo "CriticalBias: " & objItem.CriticalBias
      WScript.Echo "DefaultAlert1: " & objItem.DefaultAlert1
      WScript.Echo "DefaultAlert2: " & objItem.DefaultAlert2
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DesignedCapacity: " & objItem.DesignedCapacity
      WScript.Echo "DeviceName: " & objItem.DeviceName
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "Granularity0: " & objItem.Granularity0
      WScript.Echo "Granularity1: " & objItem.Granularity1
      WScript.Echo "Granularity2: " & objItem.Granularity2
      WScript.Echo "Granularity3: " & objItem.Granularity3
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "ManufactureDate: " & WMIDateStringToDate(objItem.ManufactureDate)
      WScript.Echo "ManufactureName: " & objItem.ManufactureName
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "SerialNumber: " & objItem.SerialNumber
      WScript.Echo "Tag: " & objItem.Tag
      WScript.Echo "Technology: " & objItem.Technology
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "UniqueID: " & objItem.UniqueID
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

