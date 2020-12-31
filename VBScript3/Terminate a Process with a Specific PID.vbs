strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where ProcessID = 2576")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
  


