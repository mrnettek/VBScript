Const SHUTDOWN = 1

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 
For Each objOperatingSystem in colOperatingSystems
    objOperatingSystem.Win32Shutdown(SHUTDOWN)
Next
  


