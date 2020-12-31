' Description: Enables per-user application of broken connection policies. To apply the same broken connection policies to all users, set the value of the BrokenConnectionPolicy property to 0 rather than 1.


Const PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    objItem.BrokenConnectionPolicy = PER_USER
    objItem.Put_
Next

