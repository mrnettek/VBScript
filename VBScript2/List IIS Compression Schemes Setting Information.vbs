' Description: Displays setting information for IIS compression schemes.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsCompressionSchemesSetting")

For Each objItem in colItems
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Hc Cache Control Header: " & _
        objItem.HcCacheControlHeader
    Wscript.Echo "Hc Compression Buffer Size: " & _
        objItem.HcCompressionBufferSize
    Wscript.Echo "Hc Compression Directory: " & _
        objItem.HcCompressionDirectory
    Wscript.Echo "Hc Do Disk Space Limiting: " & _
        objItem.HcDoDiskSpaceLimiting
    Wscript.Echo "Hc Do DynamicC ompression: " & _
        objItem.HcDoDynamicCompression
    Wscript.Echo "Hc Do On-Demand Compression: " & _
        objItem.HcDoOnDemandCompression
    Wscript.Echo "Hc Do Static Compression: " & _
        objItem.HcDoStaticCompression
    Wscript.Echo "Hc Expires Header: " & objItem.HcExpiresHeader
    Wscript.Echo "Hc Files Deleted Per Disk Free: " & _
        objItem.HcFilesDeletedPerDiskFree
    Wscript.Echo "Hc I/O Buffer Size: " & objItem.HcIoBufferSize
    Wscript.Echo "Hc Maximum Disk Space Usage: " & _
        objItem.HcMaxDiskSpaceUsage
    Wscript.Echo "Hc Maximum Queue Length: " & _
        objItem.HcMaxQueueLength
    Wscript.Echo "Hc Minimum File Size For Compression: " & _
        objItem.HcMinFileSizeForComp
    Wscript.Echo "Hc No Compression For Http 1.0: " & _
        objItem.HcNoCompressionForHttp10
    Wscript.Echo "Hc No Compression For Proxies: " & _
        objItem.HcNoCompressionForProxies
    Wscript.Echo "Hc No Compression For Range: " & _
        objItem.HcNoCompressionForRange
    Wscript.Echo "Hc Send Cache Headers: " & _
        objItem.HcSendCacheHeaders
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Setting ID: " & objItem.SettingID
Next

