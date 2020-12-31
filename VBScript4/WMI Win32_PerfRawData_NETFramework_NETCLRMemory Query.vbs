On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_NETFramework_NETCLRMemory",,48)
For Each objItem in colItems
    Wscript.Echo "AllocatedBytesPersec: " & objItem.AllocatedBytesPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FinalizationSurvivors: " & objItem.FinalizationSurvivors
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Gen0heapsize: " & objItem.Gen0heapsize
    Wscript.Echo "Gen0PromotedBytesPerSec: " & objItem.Gen0PromotedBytesPerSec
    Wscript.Echo "Gen1heapsize: " & objItem.Gen1heapsize
    Wscript.Echo "Gen1PromotedBytesPerSec: " & objItem.Gen1PromotedBytesPerSec
    Wscript.Echo "Gen2heapsize: " & objItem.Gen2heapsize
    Wscript.Echo "LargeObjectHeapsize: " & objItem.LargeObjectHeapsize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberBytesinallHeaps: " & objItem.NumberBytesinallHeaps
    Wscript.Echo "NumberGCHandles: " & objItem.NumberGCHandles
    Wscript.Echo "NumberGen0Collections: " & objItem.NumberGen0Collections
    Wscript.Echo "NumberGen1Collections: " & objItem.NumberGen1Collections
    Wscript.Echo "NumberGen2Collections: " & objItem.NumberGen2Collections
    Wscript.Echo "NumberInducedGC: " & objItem.NumberInducedGC
    Wscript.Echo "NumberofPinnedObjects: " & objItem.NumberofPinnedObjects
    Wscript.Echo "NumberofSinkBlocksinuse: " & objItem.NumberofSinkBlocksinuse
    Wscript.Echo "NumberTotalcommittedBytes: " & objItem.NumberTotalcommittedBytes
    Wscript.Echo "NumberTotalreservedBytes: " & objItem.NumberTotalreservedBytes
    Wscript.Echo "PercentTimeinGC: " & objItem.PercentTimeinGC
    Wscript.Echo "PercentTimeinGC_Base: " & objItem.PercentTimeinGC_Base
    Wscript.Echo "PromotedFinalizationMemoryfromGen0: " & objItem.PromotedFinalizationMemoryfromGen0
    Wscript.Echo "PromotedFinalizationMemoryfromGen1: " & objItem.PromotedFinalizationMemoryfromGen1
    Wscript.Echo "PromotedMemoryfromGen0: " & objItem.PromotedMemoryfromGen0
    Wscript.Echo "PromotedMemoryfromGen1: " & objItem.PromotedMemoryfromGen1
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

