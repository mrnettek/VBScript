strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")

For Each objDisk in colDisks
    Wscript.Echo objDisk.DeviceID
Next