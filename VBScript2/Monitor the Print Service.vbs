' Description: Returns the status of the Spooler service (running, stopped, paused, etc.).


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colRunningServices =  objWMIService.ExecQuery _
    ("Select * from Win32_Service Where Name = 'Spooler'")

For Each objService in colRunningServices 
    Wscript.Echo objService.DisplayName & " -- " & objService.State
Next

