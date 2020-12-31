strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("Wscript.Shell")

Do While True
    Set colProcesses = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = 'iexplore.exe'")
    If colProcesses.Count = 0 Then
        objShell.Run "iexplore.exe"
    End If
    Wscript.Sleep 60000
Loop
  


