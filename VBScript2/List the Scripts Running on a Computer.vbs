' Description: Lists the file names of all Windows Script Host scripts currently running on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = " & _
        "'Wscript.exe' OR Name = 'Cscript.exe'")
 
If colProcesses.Count = 0 Then
    Wscript.Echo "No scripts are running."
Else
    For Each objProcess in colProcesses
        Wscript.Echo objProcess.CommandLine
    Next
End If

