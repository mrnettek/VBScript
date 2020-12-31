' Description: Lists only those WMI namespaces immediately below the connected namespace.


strComputer = "."
 
Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root")

Set colNameSpaces = objWMIService.InstancesOf("__NAMESPACE")
 
For Each objNameSpace In colNameSpaces
    WScript.Echo objNameSpace.Name
Next

