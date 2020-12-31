On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_NETFramework_NETCLRExceptions",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberofExcepsThrown: " & objItem.NumberofExcepsThrown
    Wscript.Echo "NumberofExcepsThrownPersec: " & objItem.NumberofExcepsThrownPersec
    Wscript.Echo "NumberofFiltersPersec: " & objItem.NumberofFiltersPersec
    Wscript.Echo "NumberofFinallysPersec: " & objItem.NumberofFinallysPersec
    Wscript.Echo "ThrowToCatchDepthPersec: " & objItem.ThrowToCatchDepthPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

