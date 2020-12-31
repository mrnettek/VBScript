' Description: Prevents the user fabrikam\bob from being able to query Terminal Services Accounting terminal for session information.


Const WINSTATION_QUERY = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSAccount Where AccountName = 'fabrikam\\bob' AND " _
        & "TerminalName = 'Accounting'")

For Each objItem in colItems
    errResult = objItem.ModifyAuditPermissions(WINSTATION_QUERY, False)
 ext

