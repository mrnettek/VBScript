' Description: Enables per-session licensing for Terminal Services. To enable per-device licensing, pass the value 2 (rather than 4) to the ChangeMode method.


Const PER_SESSION = 4

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    errResult = objItem.ChangeMode(PER_SESSION)
Next

