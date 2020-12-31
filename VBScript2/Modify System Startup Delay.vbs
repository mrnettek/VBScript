' Description: Configures a computer to wait 10 seconds (instead of the default 30 seconds) before automatically loading the default operating system upon startup.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colStartupCommands = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")

For Each objStartupCommand in colStartupCommands
    objStartupCommand.SystemStartupDelay = 10
    objStartupCommand.Put_
Next

