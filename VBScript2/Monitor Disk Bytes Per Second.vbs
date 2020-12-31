' Description: Uses cooked performance counters to monitor disk bytes per second on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colDiskDrives = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").objectSet
objRefresher.Refresh

For i = 1 to 500
    For Each objDiskDrive in colDiskDrives
        Wscript.Echo "Drive name: " & objDiskDrive.Name
        Wscript.Echo "Disk bytes per second: " & objDiskDrive.DiskBytesPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

