' Description: Enables Terminal Services application compatibility. To disable application compatibility, set the value of the UserPermission property to 0 instead of 1.


Const ENABLED = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    objItem.UserPermission = ENABLED
    objItem.Put_
Next

