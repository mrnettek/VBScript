' Description: Deletes the Terminal Services account for fabrikam\bob. Note that in the WQL query you must separate the domain name (fabrikam) and the user name (bob) using two  slashes rather than one. Thus the account fabrikam\kenmyer would be listed as fabrikam\\kenmyer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSAccount Where AccountName = 'FABRIKAM\\bob'")

For Each objItem in colItems
    errResult = objItem.Delete()
Next

