' Description: Configures the WMI backup interval and logging level.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colWMISettings = objWMIService.ExecQuery _
    ("Select * from Win32_WMISetting")

For Each objWMISetting in colWMISettings
    objWMISetting.BackupInterval = 60
    objWMISetting.LoggingLevel = 2
    objWMISetting.Put_
Next

