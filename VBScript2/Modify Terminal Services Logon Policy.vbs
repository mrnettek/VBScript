' Description: Enables new sessions on Terminal Services. To disable new sessions, set the value of the Logons property to 0 rather than 1.


Const NEW_SESSIONS_ALLOWED = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    objItem.Logons = NEW_SESSIONS_ALLOWED
    objItem.Put_
Next

