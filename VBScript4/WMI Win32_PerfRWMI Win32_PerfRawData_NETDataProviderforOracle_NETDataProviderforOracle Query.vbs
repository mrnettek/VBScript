On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_NETDataProviderforOracle_NETDataProviderforOracle", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "HardConnectsPerSecond: " & objItem.HardConnectsPerSecond
      WScript.Echo "HardDisconnectsPerSecond: " & objItem.HardDisconnectsPerSecond
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "NumberOfActiveConnectionPoolGroups: " & objItem.NumberOfActiveConnectionPoolGroups
      WScript.Echo "NumberOfActiveConnectionPools: " & objItem.NumberOfActiveConnectionPools
      WScript.Echo "NumberOfActiveConnections: " & objItem.NumberOfActiveConnections
      WScript.Echo "NumberOfFreeConnections: " & objItem.NumberOfFreeConnections
      WScript.Echo "NumberOfInactiveConnectionPoolGroups: " & objItem.NumberOfInactiveConnectionPoolGroups
      WScript.Echo "NumberOfInactiveConnectionPools: " & objItem.NumberOfInactiveConnectionPools
      WScript.Echo "NumberOfNonPooledConnections: " & objItem.NumberOfNonPooledConnections
      WScript.Echo "NumberOfPooledConnections: " & objItem.NumberOfPooledConnections
      WScript.Echo "NumberOfReclaimedConnections: " & objItem.NumberOfReclaimedConnections
      WScript.Echo "NumberOfStasisConnections: " & objItem.NumberOfStasisConnections
      WScript.Echo "SoftConnectsPerSecond: " & objItem.SoftConnectsPerSecond
      WScript.Echo "SoftDisconnectsPerSecond: " & objItem.SoftDisconnectsPerSecond
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo
   Next
Next

