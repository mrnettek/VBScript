' Description: Restores the default permission settings for all the Terminal Service accounts on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSPermissionsSetting")

For Each objItem in colItems
    errResult = objItem.RestoreDefaults()
Next

