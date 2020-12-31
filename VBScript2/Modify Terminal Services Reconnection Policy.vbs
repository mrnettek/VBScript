' Description: Allows a user to reconnect to a Terminal Services session using any client. To require users to reconnect to a session using the same client they used previously, set the value of the ReconnectionPolicy property to 1 rather than 0.


Const PREVIOUS_CLIENT = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    objItem.ReconnectionPolicy = PREVIOUS_CLIENT
    objItem.Put_
Next

