' Description: Configures Terminal Services to use per-user connection settings. To apply the same connection settings to all users, set the value of the ConnectionPolicy property to 0 rather than 1.


Const PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    objItem.ConnectionPolicy = PER_USER
    objItem.Put_
Next

