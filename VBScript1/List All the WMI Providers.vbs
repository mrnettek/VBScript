' Description: Lists all WMI Providers installed in the root\cimv2 namespace.


strComputer = "."
 
Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2")

Set colWin32Providers = objWMIService.InstancesOf("__Win32Provider")
 
For Each objWin32Provider In colWin32Providers
    WScript.Echo objWin32Provider.Name
Next

