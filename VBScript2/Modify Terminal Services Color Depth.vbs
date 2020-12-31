' Description: Configures Terminal Services to use 16-bit color for client sessions. To use 8-bit color, pass the value 1 (rather than 3) to the SetColorDepth method. Pass the value 4 to use 24-bit color.


Const SIXTEEN_BIT_COLOR = 3
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    errResult = objItem.SetColorDepth(SIXTEEN_BIT_COLOR)
Next

