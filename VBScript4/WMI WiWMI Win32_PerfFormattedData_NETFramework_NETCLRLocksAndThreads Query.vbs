On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_NETFramework_NETCLRLocksAndThreads",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContentionRatePersec: " & objItem.ContentionRatePersec
    Wscript.Echo "CurrentQueueLength: " & objItem.CurrentQueueLength
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberofcurrentlogicalThreads: " & objItem.NumberofcurrentlogicalThreads
    Wscript.Echo "NumberofcurrentphysicalThreads: " & objItem.NumberofcurrentphysicalThreads
    Wscript.Echo "Numberofcurrentrecognizedthreads: " & objItem.Numberofcurrentrecognizedthreads
    Wscript.Echo "Numberoftotalrecognizedthreads: " & objItem.Numberoftotalrecognizedthreads
    Wscript.Echo "QueueLengthPeak: " & objItem.QueueLengthPeak
    Wscript.Echo "QueueLengthPersec: " & objItem.QueueLengthPersec
    Wscript.Echo "rateofrecognizedthreadsPersec: " & objItem.rateofrecognizedthreadsPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TotalNumberofContentions: " & objItem.TotalNumberofContentions
Next

