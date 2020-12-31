' Description: Accesses a shared folder named FinanceShare, changes the maximum number of simultaneous connections to 50, and provides a new share description.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery _
    ("Select * from Win32_Share Where Name = 'FinanceShare'")

For Each objShare in colShares
    errReturn = objShare.SetShareInfo(50, _
        "Public share for HR administrators and the Finance Group.")
Next

