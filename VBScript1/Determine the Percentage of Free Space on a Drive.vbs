strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk Where DeviceID = 'C:'")
For Each objDisk in colDisks
    Wscript.Echo objDisk.FreeSpace
Next
  


