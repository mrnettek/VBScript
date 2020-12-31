strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Winword.exe'")

For Each objProcess in colProcessList
    objProcess.GetOwner strUserName, strUserDomain 
    Wscript.Echo "Process " & objProcess.Name & " is owned by " _ 
        & strUserDomain & "\" & strUserName & "."
Next
  


