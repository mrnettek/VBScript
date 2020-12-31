On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_NETDataProviderforSqlServer_NETDataProviderforSqlServer",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "HardConnectsPerSecond: " & objItem.HardConnectsPerSecond
    Wscript.Echo "HardDisconnectsPerSecond: " & objItem.HardDisconnectsPerSecond
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfActiveConnectionPoolGroups: " & objItem.NumberOfActiveConnectionPoolGroups
    Wscript.Echo "NumberOfActiveConnectionPools: " & objItem.NumberOfActiveConnectionPools
    Wscript.Echo "NumberOfActiveConnections: " & objItem.NumberOfActiveConnections
    Wscript.Echo "NumberOfFreeConnections: " & objItem.NumberOfFreeConnections
    Wscript.Echo "NumberOfInactiveConnectionPoolGroups: " & objItem.NumberOfInactiveConnectionPoolGroups
    Wscript.Echo "NumberOfInactiveConnectionPools: " & objItem.NumberOfInactiveConnectionPools
    Wscript.Echo "NumberOfNonPooledConnections: " & objItem.NumberOfNonPooledConnections
    Wscript.Echo "NumberOfPooledConnections: " & objItem.NumberOfPooledConnections
    Wscript.Echo "NumberOfReclaimedConnections: " & objItem.NumberOfReclaimedConnections
    Wscript.Echo "NumberOfStasisConnections: " & objItem.NumberOfStasisConnections
    Wscript.Echo "SoftConnectsPerSecond: " & objItem.SoftConnectsPerSecond
    Wscript.Echo "SoftDisconnectsPerSecond: " & objItem.SoftDisconnectsPerSecond
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

