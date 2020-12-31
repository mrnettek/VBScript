' Description: Reports the account name under which each process on a computer is running.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcessList
    colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
    Wscript.Echo "Process " & objProcess.Name & " is owned by " _ 
        & strUserDomain & "\" & strNameOfUser & "."
Next

