' Description: Configures Terminal Services to use per-user time limit policies. To apply the server’s time limit policies to all users, set the value of the TimeLimitPolicy property to 0 instead of 1.


Const PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    objItem.TimeLimitPolicy = PER_USER
    objItem.Put_
Next

