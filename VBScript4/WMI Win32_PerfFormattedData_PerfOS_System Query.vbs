On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_System",,48)
For Each objItem in colItems
    Wscript.Echo "AlignmentFixupsPersec: " & objItem.AlignmentFixupsPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContextSwitchesPersec: " & objItem.ContextSwitchesPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ExceptionDispatchesPersec: " & objItem.ExceptionDispatchesPersec
    Wscript.Echo "FileControlBytesPersec: " & objItem.FileControlBytesPersec
    Wscript.Echo "FileControlOperationsPersec: " & objItem.FileControlOperationsPersec
    Wscript.Echo "FileDataOperationsPersec: " & objItem.FileDataOperationsPersec
    Wscript.Echo "FileReadBytesPersec: " & objItem.FileReadBytesPersec
    Wscript.Echo "FileReadOperationsPersec: " & objItem.FileReadOperationsPersec
    Wscript.Echo "FileWriteBytesPersec: " & objItem.FileWriteBytesPersec
    Wscript.Echo "FileWriteOperationsPersec: " & objItem.FileWriteOperationsPersec
    Wscript.Echo "FloatingEmulationsPersec: " & objItem.FloatingEmulationsPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentRegistryQuotaInUse: " & objItem.PercentRegistryQuotaInUse
    Wscript.Echo "Processes: " & objItem.Processes
    Wscript.Echo "ProcessorQueueLength: " & objItem.ProcessorQueueLength
    Wscript.Echo "SystemCallsPersec: " & objItem.SystemCallsPersec
    Wscript.Echo "SystemUpTime: " & objItem.SystemUpTime
    Wscript.Echo "Threads: " & objItem.Threads
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

