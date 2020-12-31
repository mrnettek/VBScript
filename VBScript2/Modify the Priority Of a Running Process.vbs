' Description: Changes the priority of a running instance of Notepad.exe from Normal to Above Normal.


Const ABOVE_NORMAL = 32768

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Notepad.exe'")

For Each objProcess in colProcesses
    objProcess.SetPriority(ABOVE_NORMAL) 
Next

