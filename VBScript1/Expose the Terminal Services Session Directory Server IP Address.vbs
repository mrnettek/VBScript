' Description: Allow a Terminal Services server to make its session directory IP address available. To prevent a server from exposing its session directory IP address, pass the value 0 (rather than 1) to the SetSessionDirectoryExposeServerIP method.


Const ENABLE = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSSessionDirectory")

For Each objItem in colItems
    errResult = objItem.SetSessionDirectoryExposeServerIP(ENABLE)
Next

