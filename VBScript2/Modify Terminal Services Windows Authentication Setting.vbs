' Description: Configures Terminal Services to use standard Windows authentication. To disable standard authentication (in order to use a custom authentication package), set the value of the WindowsAuthentication property to 0 rather than 1.


Const STANDARD_AUTHENTICATION = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSGeneralSetting")

For Each objItem in colItems
    objItem.WindowsAuthentication = STANDARD_AUTHENTICATION
    objItem.Put_
Next

