' Description: Uses cooked performance counters to monitor the performance of the NTFS file cache.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colCache = objRefresher.AddEnum _
    (objWMIService, "win32_PerfFormattedData_PerfOS_Cache").ObjectSet
objRefresher.Refresh

For i = 1 to 100
    For Each objCache in colCache
        Wscript.Echo "Async Copy Reads Per Second" & _
            objCache.AsyncCopyReadsPerSec
        Wscript.Echo "Async Data Maps Per Second" & _
            objCache.AsyncDataMapsPerSec
        Wscript.Echo "AsyncFastReadsPerSecond" & _
            objCache.AsyncFastReadsPerSec
        Wscript.Echo "Async MDL Reads Per Second" & _
            objCache.AsyncMDLReadsPerSec
        Wscript.Echo "Async Pin Reads Per Second" & _
            objCache.AsyncPinReadsPerSec
        Wscript.Echo "Caption" & vbTab & objCache.Caption
        Wscript.Echo "Copy Read Hits Percent " & _
            objCache.CopyReadHitsPercent
        Wscript.Echo "Copy Reads Per Second" & _
            objCache.CopyReadsPerSec
        Wscript.Echo "Data Flushes Per Second" & _
            objCache.DataFlushesPerSec
        Wscript.Echo "Data Flush Pages Per Second" & _
            objCache.DataFlushPagesPerSec
        Wscript.Echo "Data Map Hits Percent " & _
            objCache.DataMapHitsPercent
        Wscript.Echo "Data Map Pins Per Second" & _
            objCache.DataMapPinsPerSec
        Wscript.Echo "Data Maps Per Second" & _
            objCache.DataMapsPerSec
        Wscript.Echo "Description" & objCache.Description
        Wscript.Echo "Fast Read Not Possibles Per Second" & _
            objCache.FastReadNotPossiblesPerSec
        Wscript.Echo "Fast Read Resource Misses Per Second" & _
            objCache.FastReadResourceMissesPerSec
        Wscript.Echo "Fast Reads Per Second" & _
            objCache.FastReadsPerSec
        Wscript.Echo "Lazy Write Flushes Per Second" & _
            objCache.LazyWriteFlushesPerSec
        Wscript.Echo "Lazy Write Pages Per Second" & _
            objCache.LazyWritePagesPerSec
        Wscript.Echo "MDL Read Hits Percent " & _
            objCache.MDLReadHitsPercent
        Wscript.Echo "MDL Reads Per Second" & _
            objCache.MDLReadsPerSec
        Wscript.Echo "Name" & vbTab & objCache.Name
        Wscript.Echo "Pin Read Hits Percent" & _
            objCache.PinReadHitsPercent
        Wscript.Echo "Pin Reads Per Second" & _
            objCache.PinReadsPerSec
        Wscript.Echo "Read Aheads Per Second" & _
            objCache.ReadAheadsPerSec
        Wscript.Echo "Sync Copy Reads Per Second" & _
            objCache.SyncCopyReadsPerSec
        Wscript.Echo "Sync Data Maps Per Second" & _
            objCache.SyncDataMapsPerSec
        Wscript.Echo "Sync Fast Reads Per Second" & _
            objCache.SyncFastReadsPerSec
        Wscript.Echo "Sync MDL Reads Per Second" & _
            objCache.SyncMDLReadsPerSec
        Wscript.Echo "Sync Pin Reads Per Second" & _
            objCache.SyncPinReadsPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

