Const BELOW_NORMAL = 16384

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Notepad.exe'")
For Each objProcess in colProcesses
    objProcess.SetPriority(BELOW_NORMAL) 
Next
  


