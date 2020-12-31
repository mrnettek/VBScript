' Description: Sets the auto-delay time for Autochk.exe to 30 seconds.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colAutoChkSettings = objWMIService.ExecQuery _
    ("Select * from Win32_AutochkSetting")

For Each objAutoChkSetting in colAutoChkSettings
    objAutoChkSetting.UserInputDelay = 30
    objAutoChkSetting.Put_
Next

