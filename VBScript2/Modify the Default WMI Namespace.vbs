' Description: Sets the WMI "Default namespace for scripting" setting to "root\cimv2".


strComputer = "."
 
Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2")

Set colWMISettings = objWMIService.InstancesOf("Win32_WMISetting")
 
For Each objWMISetting in colWMISettings
    objWMISetting.ASPScriptDefaultNamespace = "root\cimv2"
    objWMISetting.Put_
Next

