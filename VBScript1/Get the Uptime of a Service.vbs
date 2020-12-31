strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
    ("Select * From Win32_Service Where Name = 'Spooler'")
 
For Each objService in colServices
    intProcessID = objService.ProcessID

    Set colProcesses = objWMIService.ExecQuery _
        ("Select * From Win32_Process Where ProcessID = " & intProcessID)

    For Each objProcess in colProcesses
        Wscript.Echo objProcess.CreationDate
    Next
Next
  


