' Description: Returns information about the compressions schemes on an IIS server.


strComputer = "LocalHost"
Set objIIS = GetObject _
    ("IIS://" & strComputer & "/W3SVC/Filters/Compression/Parameters")
 
Wscript.Echo "Cache Control Header: " & objIIS.HcCacheControlHeader
Wscript.Echo "Compression Buffer Size: " & objIIS.HcCompressionBufferSize
Wscript.Echo "Compression Directory: " & objIIS.HcCompressionDirectory
Wscript.Echo "Do Disk Space Limiting: " & objIIS.HcDoDiskSpaceLimiting
Wscript.Echo "Do Dynamic Compression: " & objIIS.HcDoDynamicCompression
Wscript.Echo "Do On-Demand Compression: " & objIIS.HcDoOnDemandCompression
Wscript.Echo "Do Static Compression: " & objIIS.HcDoStaticCompression
Wscript.Echo "Expires Header: " & objIIS.HcExpiresHeader
Wscript.Echo "Files Deleted Per Disk Free: " &  _
    objIIS.HcFilesDeletedPerDiskFree
Wscript.Echo "I/O Buffer Size: " & objIIS.HcIoBufferSize
Wscript.Echo "Maximum Disk Space Usage: " & objIIS.HcMaxDiskSpaceUsage
Wscript.Echo "Maximum Queue Length: " & objIIS.HcMaxQueueLength
Wscript.Echo "Minimum File Size for Compression: " &  _
    objIIS.HcMinFileSizeForComp
Wscript.Echo "No Compression for HTTP 1.0: " &  _
    objIIS.HcNoCompressionForHttp10
Wscript.Echo "No Compression for Proxies: " &  _
    objIIS.HcNoCompressionForProxies
Wscript.Echo "No Compression for Range: " &  _
    objIIS.HcNoCompressionForRange
Wscript.Echo "Send Cache Headers: " & objIIS.HcSendCacheHeaders

