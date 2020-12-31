' Description: Identifies the file system in use for each logical disk on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")

For Each objDisk in colDisks
    Wscript.Echo "Device ID: "& vbTab &  objDisk.DeviceID       
    Wscript.Echo "File System: "& vbTab & objDisk.FileSystem
Next

