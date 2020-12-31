' Description: Enables per-user application of client logon policies. To apply the same policies to all users, set the value of the ClientLogonInfoPolicy property to 0 rather than 1.


Const PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSLogonSetting")

For Each objItem in colItems
    objItem.ClientLogonInfoPolicy = PER_USER
    objItem.Put_
Next

