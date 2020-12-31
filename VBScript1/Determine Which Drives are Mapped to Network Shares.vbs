strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colDrives = objWMIService.ExecQuery _
    ("Select * From Win32_LogicalDisk Where DriveType = 4")

For Each objDrive in colDrives
    Wscript.Echo "Drive letter: " & objDrive.DeviceID
    Wscript.Echo "Network path: " & objDrive.ProviderName
Next
  


