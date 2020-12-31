' Description: Enables the color depth policy on a computer running Terminal Services. To disable this policy, pass the value 0 (rather than 1) to the SetColorDepthPolicy method.


Const ENABLE = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    errResult = objItem.SetColorDepthPolicy(ENABLE)
Next

