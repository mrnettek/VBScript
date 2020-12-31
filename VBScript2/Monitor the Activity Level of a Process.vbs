strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

For i = 1 to 11

    Set colProcesses = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = 'Notepad.exe'")

    For Each objProcess in colProcesses
        sngProcessTime = (CSng(objProcess.KernelModeTime) + _
                CSng(objProcess.UserModeTime)) / 10000000
        Wscript.Echo objProcess.Name, sngProcessTime
    Next

    Wscript.Sleep 30000
Next
  


