On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
For Each objItem in colItems
    Wscript.Echo "AvailableBytes: " & objItem.AvailableBytes
    Wscript.Echo "AvailableKBytes: " & objItem.AvailableKBytes
    Wscript.Echo "AvailableMBytes: " & objItem.AvailableMBytes
    Wscript.Echo "CacheBytes: " & objItem.CacheBytes
    Wscript.Echo "CacheBytesPeak: " & objItem.CacheBytesPeak
    Wscript.Echo "CacheFaultsPersec: " & objItem.CacheFaultsPersec
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CommitLimit: " & objItem.CommitLimit
    Wscript.Echo "CommittedBytes: " & objItem.CommittedBytes
    Wscript.Echo "DemandZeroFaultsPersec: " & objItem.DemandZeroFaultsPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FreeSystemPageTableEntries: " & objItem.FreeSystemPageTableEntries
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PageFaultsPersec: " & objItem.PageFaultsPersec
    Wscript.Echo "PageReadsPersec: " & objItem.PageReadsPersec
    Wscript.Echo "PagesInputPersec: " & objItem.PagesInputPersec
    Wscript.Echo "PagesOutputPersec: " & objItem.PagesOutputPersec
    Wscript.Echo "PagesPersec: " & objItem.PagesPersec
    Wscript.Echo "PageWritesPersec: " & objItem.PageWritesPersec
    Wscript.Echo "PercentCommittedBytesInUse: " & objItem.PercentCommittedBytesInUse
    Wscript.Echo "PoolNonpagedAllocs: " & objItem.PoolNonpagedAllocs
    Wscript.Echo "PoolNonpagedBytes: " & objItem.PoolNonpagedBytes
    Wscript.Echo "PoolPagedAllocs: " & objItem.PoolPagedAllocs
    Wscript.Echo "PoolPagedBytes: " & objItem.PoolPagedBytes
    Wscript.Echo "PoolPagedResidentBytes: " & objItem.PoolPagedResidentBytes
    Wscript.Echo "SystemCacheResidentBytes: " & objItem.SystemCacheResidentBytes
    Wscript.Echo "SystemCodeResidentBytes: " & objItem.SystemCodeResidentBytes
    Wscript.Echo "SystemCodeTotalBytes: " & objItem.SystemCodeTotalBytes
    Wscript.Echo "SystemDriverResidentBytes: " & objItem.SystemDriverResidentBytes
    Wscript.Echo "SystemDriverTotalBytes: " & objItem.SystemDriverTotalBytes
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TransitionFaultsPersec: " & objItem.TransitionFaultsPersec
    Wscript.Echo "WriteCopiesPersec: " & objItem.WriteCopiesPersec
Next

