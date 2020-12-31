' Description: Returns a list of all the services installed on a computer, and indicates their current status (typically, running or not running).


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colRunningServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colRunningServices 
    Wscript.Echo objService.DisplayName  & VbTab & objService.State
Next

