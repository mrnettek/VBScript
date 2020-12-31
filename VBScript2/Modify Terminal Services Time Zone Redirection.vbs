' Description: Disables Terminal Service time zone redirection on a computer. To enable time zone redirection, pass the value 1 (instead of 0) to the SetTimeZoneRedirection method.


Const DISABLE = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    errResult = objItem.SetTimeZoneRedirection(DISABLE)
Next

