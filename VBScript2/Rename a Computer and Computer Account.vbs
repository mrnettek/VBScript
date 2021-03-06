' Description: Renames a computer and its corresponding Active Directory computer account. Requires Windows XP or Windows Server 2003, and must be run on the local computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colComputers = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")

For Each objComputer in colComputers
    err = objComputer.Rename("WebServer")
Next

