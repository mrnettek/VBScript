' Description: Modifies the descriptive comment attached to a Terminal Services server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSGeneralSetting")

For Each objItem in colItems
    objItem.Comment = "Accounting session."
    objItem.Put_
Next

