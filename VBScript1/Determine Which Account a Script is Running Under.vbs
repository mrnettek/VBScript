strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where " & _
    "Name = 'cscript.exe' or Name = 'wscript.exe'")

For Each objProcess in colProcessList
    If InStr(objProcess.CommandLine, "test.vbs") Then
        colProperties =   objProcess.GetOwner(strNameOfUser,strUserDomain)
        Wscript.Echo "This script is running under the account belonging to " _ 
            & strUserDomain & "\" & strNameOfUser & "."
    End If
Next
  


