On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Spooler_PrintQueue",,48)
For Each objItem in colItems
    Wscript.Echo "AddNetworkPrinterCalls: " & objItem.AddNetworkPrinterCalls
    Wscript.Echo "BytesPrintedPersec: " & objItem.BytesPrintedPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "EnumerateNetworkPrinterCalls: " & objItem.EnumerateNetworkPrinterCalls
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "JobErrors: " & objItem.JobErrors
    Wscript.Echo "Jobs: " & objItem.Jobs
    Wscript.Echo "JobsSpooling: " & objItem.JobsSpooling
    Wscript.Echo "MaxJobsSpooling: " & objItem.MaxJobsSpooling
    Wscript.Echo "MaxReferences: " & objItem.MaxReferences
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NotReadyErrors: " & objItem.NotReadyErrors
    Wscript.Echo "OutofPaperErrors: " & objItem.OutofPaperErrors
    Wscript.Echo "References: " & objItem.References
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TotalJobsPrinted: " & objItem.TotalJobsPrinted
    Wscript.Echo "TotalPagesPrinted: " & objItem.TotalPagesPrinted
Next

