' Description: Uses cooked performance counters to monitor performance of the logical disk drives installed on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colDisks = objRefresher.AddEnum _
    (objWMIService, "win32_perfformatteddata_perfdisk_logicaldisk"). _
        objectSet
objRefresher.Refresh

For i = 1 to 100
    For Each objDisk in colDisks
        Wscript.Echo "Average Disk Bytes Per Read: " & _
            objDisk.AvgDiskBytesPerRead
        Wscript.Echo "Average Disk Bytes Per Transfer: " & _
            objDisk.AvgDiskBytesPerTransfer
        Wscript.Echo "Average Disk Bytes Per Write: " & _
            objDisk.AvgDiskBytesPerWrite
        Wscript.Echo "Average Disk Queue Length: " & _
            objDisk.AvgDiskQueueLength
        Wscript.Echo "Average Disk Read Queue Length: " & _
            objDisk.AvgDiskReadQueueLength
        Wscript.Echo "Average Disk Seconds Per Read: " & _
            objDisk.AvgDiskSecPerRead
        Wscript.Echo "Average Disk Seconds Per Transfer: " & _
            objDisk.AvgDiskSecPerTransfer
        Wscript.Echo "Average Disk Seconds Per Write: " & _
            objDisk.AvgDiskSecPerWrite
        Wscript.Echo "Average Disk Write Queue Length: " & _
            objDisk.AvgDiskWriteQueueLength
        Wscript.Echo "Current Disk Queue Length: " & _
            objDisk.CurrentDiskQueueLength
        Wscript.Echo "Disk Bytes Per Second: " & _
            objDisk.DiskBytesPerSec
        Wscript.Echo "Disk Read Bytes Per Second: " & _
            objDisk.DiskReadBytesPerSec
        Wscript.Echo "Disk Reads Per Second: " & _
            objDisk.DiskReadsPerSec
        Wscript.Echo "Disk Transfers Per Second: " & _
            objDisk.DiskTransfersPerSec
        Wscript.Echo "Disk Write Bytes Per Second: " & _
            objDisk.DiskWriteBytesPerSec
        Wscript.Echo "Disk Writes Per Second: " & _
            objDisk.DiskWritesPerSec
        Wscript.Echo "Free Megabytes: " & objDisk.FreeMegabytes
        Wscript.Echo "Name: " & objDisk.Name
        Wscript.Echo "Percent Disk Read Time: " &  _
            objDisk.PercentDiskReadTime
        Wscript.Echo "Percent Disk Time: " & _
            objDisk.PercentDiskTime
        Wscript.Echo "Percent Disk Write Time: " & _
            objDisk.PercentDiskWriteTime
        Wscript.Echo "Percent Free Space: " & _
            objDisk.PercentFreeSpace
        Wscript.Echo "Percent Idle Time: " & _
            objDisk.PercentIdleTime
        Wscript.Echo "Split IO Per Second: " & _
            objDisk.SplitIOPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

