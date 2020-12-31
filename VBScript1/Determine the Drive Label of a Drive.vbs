strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_LogicalDisk Where DeviceID = 'D:'")

For Each objItem in colItems
    Wscript.Echo objItem.VolumeName 
Next
  


