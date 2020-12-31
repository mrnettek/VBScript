On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_JobObject",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CurrentPercentKernelModeTime: " & objItem.CurrentPercentKernelModeTime
    Wscript.Echo "CurrentPercentProcessorTime: " & objItem.CurrentPercentProcessorTime
    Wscript.Echo "CurrentPercentUserModeTime: " & objItem.CurrentPercentUserModeTime
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PagesPerSec: " & objItem.PagesPerSec
    Wscript.Echo "ProcessCountActive: " & objItem.ProcessCountActive
    Wscript.Echo "ProcessCountTerminated: " & objItem.ProcessCountTerminated
    Wscript.Echo "ProcessCountTotal: " & objItem.ProcessCountTotal
    Wscript.Echo "ThisPeriodmSecKernelMode: " & objItem.ThisPeriodmSecKernelMode
    Wscript.Echo "ThisPeriodmSecProcessor: " & objItem.ThisPeriodmSecProcessor
    Wscript.Echo "ThisPeriodmSecUserMode: " & objItem.ThisPeriodmSecUserMode
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TotalmSecKernelMode: " & objItem.TotalmSecKernelMode
    Wscript.Echo "TotalmSecProcessor: " & objItem.TotalmSecProcessor
    Wscript.Echo "TotalmSecUserMode: " & objItem.TotalmSecUserMode
Next

