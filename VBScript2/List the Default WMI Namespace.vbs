' Description: Retrieves and displays the current WMI "Default namespace for scripting" setting.


strComputer = "."
 
Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2")

Set colWMISettings = objWMIService.InstancesOf("Win32_WMISetting")
 
For Each objWMISetting in colWMISettings
    Wscript.Echo "Default namespace for scripting: " & _
    objWMISetting.ASPScriptDefaultNamespace 
Next

