' Description: Enable the session directory on a Terminal Services server. To disable the session directory, pass the value 0 (rather than 1) to the SetSessionDirectoryActive method.


Const ENABLE_SESSION_DIRECTORY = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSSessionDirectory")

For Each objItem in colItems
    errResult = objItem.SetSessionDirectoryActive(ENABLE_SESSION_DIRECTORY)
Next

