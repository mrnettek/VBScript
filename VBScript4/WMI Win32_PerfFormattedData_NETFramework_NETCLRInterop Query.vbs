On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_NETFramework_NETCLRInterop",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberofCCWs: " & objItem.NumberofCCWs
    Wscript.Echo "Numberofmarshalling: " & objItem.Numberofmarshalling
    Wscript.Echo "NumberofStubs: " & objItem.NumberofStubs
    Wscript.Echo "NumberofTLBexportsPersec: " & objItem.NumberofTLBexportsPersec
    Wscript.Echo "NumberofTLBimportsPersec: " & objItem.NumberofTLBimportsPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

