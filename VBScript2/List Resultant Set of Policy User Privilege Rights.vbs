' Description: Returns information about user privileges assigned by using Group Policy.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_UserPrivilegeRight")

For Each objItem in colItems
    For Each strAccountList in objItem.AccountList
        Wscript.Echo "Account list: " & strAccountList
    Next
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "User Right: " & objItem.UserRight
    Wscript.Echo
Next

