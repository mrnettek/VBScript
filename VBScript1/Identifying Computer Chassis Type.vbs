' Description: Retrieves an integer value indicating the chassis type for a computer (mini-tower, laptop, etc.). The script does not include a description of each value that can be returned.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colChassis = objWMIService.ExecQuery _
    ("Select * from Win32_SystemEnclosure")

For Each objChassis in colChassis
    For i = Lbound(objChassis.ChassisTypes) to Ubound(objChassis.ChassisTypes)
        Wscript.Echo objChassis.ChassisTypes(i)
    Next
Next

