strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colAccounts = objWMIService.ExecQuery _
    ("Select * From Win32_Group Where LocalAccount = TRUE And SID = 'S-1-5-32-544'")

For Each objAccount in colAccounts
    Wscript.Echo objAccount.Name
Next
  


