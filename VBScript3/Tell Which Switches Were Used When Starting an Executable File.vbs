strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Process Where Name = 'netstat.exe'")

For Each objItem in colItems
    Wscript.Echo objItem.CommandLine
Next
  


