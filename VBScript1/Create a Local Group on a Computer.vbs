' Description: Creates a local group named FinanceUsers on a computer named atl-ws-01.


strComputer = "atl-ws-01"
Set colAccounts = GetObject("WinNT://" & strComputer & "")
Set objUser = colAccounts.Create("group", "FinanceUsers")
objUser.SetInfo

