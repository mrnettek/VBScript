' Description: Uses cooked performance counters to retrieve free disk space on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDiskDrives = objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_PerfDisk_LogicalDisk Where " _
        & "Name <> '_Total'")

For Each objDiskDrive in colDiskDrives
    Wscript.Echo "Drive Name: " & objDiskDrive.Name
    Wscript.Echo "Free Space: " & objDiskDrive.FreeMegabytes
Next

