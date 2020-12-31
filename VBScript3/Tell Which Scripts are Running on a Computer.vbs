strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'cscript.exe'" & _
        " OR Name = 'wscript.exe'")
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next
  


