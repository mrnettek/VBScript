' Description: Uses cooked performance counters to monitor Terminal Service session performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_Perf_TermService_TerminalServiceSession").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Handle Count: " & objItem.HandleCount
        Wscript.Echo "Input Asynchronous Frame Error: " & _
            objItem.InputAsyncFrameError
        Wscript.Echo "Input Asynchronous Overflow: " & _
            objItem.InputAsyncOverflow
        Wscript.Echo "Input Asynchronous Overrun: " & objItem.InputAsyncOverrun
        Wscript.Echo "Input Asynchronous Parity Error: " & _
            objItem.InputAsyncParityError
        Wscript.Echo "Input Bytes: " & objItem.InputBytes
        Wscript.Echo "Input Compressed Bytes: " & objItem.InputCompressedBytes
        Wscript.Echo "Input Compress Flushes: " & objItem.InputCompressFlushes
        Wscript.Echo "Input Compression Ratio: " & _
            objItem.InputCompressionRatio
        Wscript.Echo "Input Errors: " & objItem.InputErrors
        Wscript.Echo "Input Frames: " & objItem.InputFrames
        Wscript.Echo "Input Timeouts: " & objItem.InputTimeouts
        Wscript.Echo "Input Transport Errors: " & objItem.InputTransportErrors
        Wscript.Echo "Input Wait For OutputBuffer: " & _
            objItem.InputWaitForOutBuf
        Wscript.Echo "Input Wd Bytes: " & objItem.InputWdBytes
        Wscript.Echo "Input Wd Frames: " & objItem.InputWdFrames
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Output Asynchronous Frame Error: " & _
            objItem.OutputAsyncFrameError
        Wscript.Echo "Output Asynchronous Overflow: " & _
            objItem.OutputAsyncOverflow
        Wscript.Echo "Output Asynchronous Overrun: " & _
            objItem.OutputAsyncOverrun
        Wscript.Echo "OutputAsynchronous Parity Error: " & _
            objItem.OutputAsyncParityError
        Wscript.Echo "Output Bytes: " & objItem.OutputBytes
        Wscript.Echo "Output Compressed Bytes: " & _
            objItem.OutputCompressedBytes
        Wscript.Echo "Output Compress Flushes: " & _
            objItem.OutputCompressFlushes
        Wscript.Echo "Output Compression Ratio: " & _
            objItem.OutputCompressionRatio
        Wscript.Echo "Output Errors: " & objItem.OutputErrors
        Wscript.Echo "Output Frames: " & objItem.OutputFrames
        Wscript.Echo "Output Timeouts: " & objItem.OutputTimeouts
        Wscript.Echo "Output Transport Errors: " & _
            objItem.OutputTransportErrors
        Wscript.Echo "Output Wait For Outout Buffer: " & _
            objItem.OutputWaitForOutBuf
        Wscript.Echo "Output Wd Bytes: " & objItem.OutputWdBytes
        Wscript.Echo "Output Wd Frames: " & objItem.OutputWdFrames
        Wscript.Echo "Page Faults Per Second: " & objItem.PageFaultsPersec
        Wscript.Echo "Page File Bytes: " & objItem.PageFileBytes
        Wscript.Echo "Page File Bytes Peak: " & objItem.PageFileBytesPeak
        Wscript.Echo "Percent Privileged Time: " & _
            objItem.PercentPrivilegedTime
        Wscript.Echo "Percent Processor Time: " & objItem.PercentProcessorTime
        Wscript.Echo "Percent User Time: " & objItem.PercentUserTime
        Wscript.Echo "Pool Nonpaged Bytes: " & objItem.PoolNonpagedBytes
        Wscript.Echo "Pool Paged Bytes: " & objItem.PoolPagedBytes
        Wscript.Echo "Private Bytes: " & objItem.PrivateBytes
        Wscript.Echo "Protocol Bitmap Cache Hit Ratio: " & _
            objItem.ProtocolBitmapCacheHitRatio
        Wscript.Echo "Protocol Bitmap Cache Hits: " & _
            objItem.ProtocolBitmapCacheHits
        Wscript.Echo "Protocol Bitmap Cache Reads: " & _
            objItem.ProtocolBitmapCacheReads
        Wscript.Echo "Protocol Brush Cache Hit Ratio: " & _
            objItem.ProtocolBrushCacheHitRatio
        Wscript.Echo "Protocol Brush Cache Hits: " & _
            objItem.ProtocolBrushCacheHits
        Wscript.Echo "Protocol Brush Cache Reads: " & _
            objItem.ProtocolBrushCacheReads
        Wscript.Echo "Protocol Glyph Cache Hit Ratio: " & _
            objItem.ProtocolGlyphCacheHitRatio
        Wscript.Echo "Protocol Glyph Cache Hits: " & _
            objItem.ProtocolGlyphCacheHits
        Wscript.Echo "Protocol Glyph Cache Reads: " & _)
            objItem.ProtocolGlyphCacheReads
        Wscript.Echo "Protocol Save Screen Bitmap Cache Hit Ratio: " & _
            objItem.ProtocolSaveScreenBitmapCacheHitRatio
        Wscript.Echo "Protocol Save Screen Bitmap Cache Hits: " & _
            objItem.ProtocolSaveScreenBitmapCacheHits
        Wscript.Echo "Protocol Save Screen Bitmap Cache Reads: " & _
            objItem.ProtocolSaveScreenBitmapCacheReads
        Wscript.Echo "Thread Count: " & objItem.ThreadCount
        Wscript.Echo "Total Asynchronous Frame Error: " & _
            objItem.TotalAsyncFrameError
        Wscript.Echo "Total Asynchronous Overflow: " & _
            objItem.TotalAsyncOverflow
        Wscript.Echo "Total Asynchronous Overrun: " & objItem.TotalAsyncOverrun
        Wscript.Echo "Total Asynchronous Parity Error: " & _
            objItem.TotalAsyncParityError
        Wscript.Echo "Total Bytes: " & objItem.TotalBytes
        Wscript.Echo "Total Compressed Bytes: " & objItem.TotalCompressedBytes
        Wscript.Echo "Total Compress Flushes: " & objItem.TotalCompressFlushes
        Wscript.Echo "Total Compression Ratio: " & _
            objItem.TotalCompressionRatio
        Wscript.Echo "Total Errors: " & objItem.TotalErrors
        Wscript.Echo "Total Frames: " & objItem.TotalFrames
        Wscript.Echo "Total Protocol Cache Hit Ratio: " & _
            objItem.TotalProtocolCacheHitRatio
        Wscript.Echo "Total Protocol Cache Hits: " & _
            objItem.TotalProtocolCacheHits
        Wscript.Echo "Total Protocol Cache Reads: " & _
            objItem.TotalProtocolCacheReads
        Wscript.Echo "Total Timeouts: " & objItem.TotalTimeouts
        Wscript.Echo "Total Transport Errors: " & objItem.TotalTransportErrors
        Wscript.Echo "Total Wait For Output Buffer: " & _
            objItem.TotalWaitForOutBuf
        Wscript.Echo "Total Wd Bytes: " & objItem.TotalWdBytes
        Wscript.Echo "Total Wd Frames: " & objItem.TotalWdFrames
        Wscript.Echo "Virtual Bytes: " & objItem.VirtualBytes
        Wscript.Echo "Virtual Bytes Peak: " & objItem.VirtualBytesPeak
        Wscript.Echo "Working Set: " & objItem.WorkingSet
        Wscript.Echo "Working Set Peak: " & objItem.WorkingSetPeak
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

