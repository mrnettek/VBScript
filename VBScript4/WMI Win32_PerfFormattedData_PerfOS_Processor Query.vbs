On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Processor",,48)
For Each objItem in colItems
    Wscript.Echo "C1TransitionsPersec: " & objItem.C1TransitionsPersec
    Wscript.Echo "C2TransitionsPersec: " & objItem.C2TransitionsPersec
    Wscript.Echo "C3TransitionsPersec: " & objItem.C3TransitionsPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DPCRate: " & objItem.DPCRate
    Wscript.Echo "DPCsQueuedPersec: " & objItem.DPCsQueuedPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "InterruptsPersec: " & objItem.InterruptsPersec
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentC1Time: " & objItem.PercentC1Time
    Wscript.Echo "PercentC2Time: " & objItem.PercentC2Time
    Wscript.Echo "PercentC3Time: " & objItem.PercentC3Time
    Wscript.Echo "PercentDPCTime: " & objItem.PercentDPCTime
    Wscript.Echo "PercentIdleTime: " & objItem.PercentIdleTime
    Wscript.Echo "PercentInterruptTime: " & objItem.PercentInterruptTime
    Wscript.Echo "PercentPrivilegedTime: " & objItem.PercentPrivilegedTime
    Wscript.Echo "PercentProcessorTime: " & objItem.PercentProcessorTime
    Wscript.Echo "PercentUserTime: " & objItem.PercentUserTime
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

