' Description: Sets the session directory location of a Terminal Services server to 192.168.1.3.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSSessionDirectory")

For Each objItem in colItems
    errResult = objItem.SetSessionDirectoryProperty _
        ("SessionDirectoryLocation", "192.168.1.3")
Next

