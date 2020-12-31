On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfOS_Cache",,48)
For Each objItem in colItems
    Wscript.Echo "AsyncCopyReadsPersec: " & objItem.AsyncCopyReadsPersec
    Wscript.Echo "AsyncDataMapsPersec: " & objItem.AsyncDataMapsPersec
    Wscript.Echo "AsyncFastReadsPersec: " & objItem.AsyncFastReadsPersec
    Wscript.Echo "AsyncMDLReadsPersec: " & objItem.AsyncMDLReadsPersec
    Wscript.Echo "AsyncPinReadsPersec: " & objItem.AsyncPinReadsPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CopyReadHitsPercent: " & objItem.CopyReadHitsPercent
    Wscript.Echo "CopyReadHitsPercent_Base: " & objItem.CopyReadHitsPercent_Base
    Wscript.Echo "CopyReadsPersec: " & objItem.CopyReadsPersec
    Wscript.Echo "DataFlushesPersec: " & objItem.DataFlushesPersec
    Wscript.Echo "DataFlushPagesPersec: " & objItem.DataFlushPagesPersec
    Wscript.Echo "DataMapHitsPercent: " & objItem.DataMapHitsPercent
    Wscript.Echo "DataMapHitsPercent_Base: " & objItem.DataMapHitsPercent_Base
    Wscript.Echo "DataMapPinsPersec: " & objItem.DataMapPinsPersec
    Wscript.Echo "DataMapPinsPersec_Base: " & objItem.DataMapPinsPersec_Base
    Wscript.Echo "DataMapsPersec: " & objItem.DataMapsPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FastReadNotPossiblesPersec: " & objItem.FastReadNotPossiblesPersec
    Wscript.Echo "FastReadResourceMissesPersec: " & objItem.FastReadResourceMissesPersec
    Wscript.Echo "FastReadsPersec: " & objItem.FastReadsPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "LazyWriteFlushesPersec: " & objItem.LazyWriteFlushesPersec
    Wscript.Echo "LazyWritePagesPersec: " & objItem.LazyWritePagesPersec
    Wscript.Echo "MDLReadHitsPercent: " & objItem.MDLReadHitsPercent
    Wscript.Echo "MDLReadHitsPercent_Base: " & objItem.MDLReadHitsPercent_Base
    Wscript.Echo "MDLReadsPersec: " & objItem.MDLReadsPersec
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PinReadHitsPercent: " & objItem.PinReadHitsPercent
    Wscript.Echo "PinReadHitsPercent_Base: " & objItem.PinReadHitsPercent_Base
    Wscript.Echo "PinReadsPersec: " & objItem.PinReadsPersec
    Wscript.Echo "ReadAheadsPersec: " & objItem.ReadAheadsPersec
    Wscript.Echo "SyncCopyReadsPersec: " & objItem.SyncCopyReadsPersec
    Wscript.Echo "SyncDataMapsPersec: " & objItem.SyncDataMapsPersec
    Wscript.Echo "SyncFastReadsPersec: " & objItem.SyncFastReadsPersec
    Wscript.Echo "SyncMDLReadsPersec: " & objItem.SyncMDLReadsPersec
    Wscript.Echo "SyncPinReadsPersec: " & objItem.SyncPinReadsPersec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

