strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_CDROMDrive Where VolumeName = 'MyApps'")

Wscript.Echo colItems.Count
  


