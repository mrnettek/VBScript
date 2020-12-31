' Description: Gives the user fabrikam\bob the right to connect to another Terminal Services session.


CONST WINSTATION_CONNECT = 8 
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSAccount Where AccountName = 'fabrikam\\bob' " & _
        "AND TerminalName = 'Accounting'")

For Each objItem in colItems
    errResult = objItem.ModifyPermissions(WINSTATION_CONNECT,True)
Next

