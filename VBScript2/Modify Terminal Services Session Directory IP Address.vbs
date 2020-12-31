' Description: Configures Terminal Services to use a session directory with the IP address of 192.168.1.3.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSSessionDirectory")

For Each objItem in colItems
    objItem.SessionDirectoryIPAddress = "192.168.1.3"
    objItem.Put_
Next

