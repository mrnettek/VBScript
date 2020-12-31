' Description: Enables Terminal Services to apply per-user remote control policies. To apply the same remote control policies to all users, set the value of the RemoteControlPolicy property to 0 rather than 1.


Const PER_USER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSRemoteControlSetting")

For Each objItem in colItems
    objItem.RemoteControlPolicy = PER_USER
    objItem.Put_
Next

