strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'notepad.exe'")

For Each objProcess in colProcessList
    objProcess.GetOwner strNameOfUser,strUserDomain
    strOwner = strUserDomain & "\" & strNameOfUser
    If LCase(strOwner) = "fabrikam\kenmyer" Then
        Wscript.Echo objProcess.Handle
    End If
Next
  


