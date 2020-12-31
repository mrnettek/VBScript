' Description: Uses cooked performance counters to monitor physical disk performance.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfDisk_PhysicalDisk").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Average Disk Bytes Per Read: " & _
            objItem.AvgDiskBytesPerRead
        Wscript.Echo "Average Disk Bytes Per Transfer: " & _
            objItem.AvgDiskBytesPerTransfer
        Wscript.Echo "Average Disk Bytes Per Write: " & _
            objItem.AvgDiskBytesPerWrite
        Wscript.Echo "Average Disk Queue Length: " & objItem.AvgDiskQueueLength
        Wscript.Echo "Average Disk Read Queue Length: " & _
            objItem.AvgDiskReadQueueLength
        Wscript.Echo "Average Disk Seconds Per Read: " & _
            objItem.AvgDisksecPerRead
        Wscript.Echo "Average Disk Seconds Per Transfer: " & _
            objItem.AvgDisksecPerTransfer
        Wscript.Echo "Average Disk Seconds Per Write: " & _
            objItem.AvgDisksecPerWrite
        Wscript.Echo "Average Disk Write Queue Length: " & _
            objItem.AvgDiskWriteQueueLength
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Current Disk Queue Length: " & _
            objItem.CurrentDiskQueueLength
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Disk Bytes Per Second: " & objItem.DiskBytesPersec
        Wscript.Echo "Disk Read Bytes Per Second: " & _
            objItem.DiskReadBytesPersec
        Wscript.Echo "Disk Reads Per Second: " & objItem.DiskReadsPersec
        Wscript.Echo "Disk Transfers Per Second: " & _
            objItem.DiskTransfersPersec
        Wscript.Echo "Disk Write Bytes Per Second: " & _
            objItem.DiskWriteBytesPersec
        Wscript.Echo "Disk Writes Per Second: " & objItem.DiskWritesPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Disk Read Time: " & objItem.PercentDiskReadTime
        Wscript.Echo "Percent Disk Time: " & objItem.PercentDiskTime
        Wscript.Echo "Percent Disk Write Time: " & objItem.PercentDiskWriteTime
        Wscript.Echo "Percent Idle Time: " & objItem.PercentIdleTime
        Wscript.Echo "Split I/O Per Second: " & objItem.SplitIOPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

