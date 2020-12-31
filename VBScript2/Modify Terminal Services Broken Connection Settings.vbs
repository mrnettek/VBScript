' Description: Configures Terminal Services to permanently delete a session in case of a broken connection. To configure Terminal Services to merely disconnect the user from the session instead, pass the value 0 (instead of 1) to the BrokenConnection method.


Const DISCONNECT_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    errResult = objItem.BrokenConnection(DISCONNECT_USER)
Next

