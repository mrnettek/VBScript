' Description: Configures Terminal Services to assign startup programs on a per-user basis. To have Terminal Services apply the same startup program to all users, set the value of the InitialProgramPolicy property to 0 rather than 1.


CONST PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSEnvironmentSetting")

For Each objItem in colItems
    objItem.InitialProgramPolicy = PER_USER
    objItem.Put_
Next

