' Description: Changes the volume label for drive C to "Finance Volume."


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDrives = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk where DeviceID = 'C:'")

For Each objDrive in colDrives
    objDrive.VolumeName = "Finance Volume"
    objDrive.Put_
Next

