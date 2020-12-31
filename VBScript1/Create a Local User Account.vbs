' Description: Creates a local user account (Admin2) on a computer named atl-ws-01, and sets the password for the account to 09iu%4et.


strComputer = "atl-ws-01"
Set colAccounts = GetObject("WinNT://" & strComputer & "")
Set objUser = colAccounts.Create("user", "Admin2")
objUser.SetPassword "09iu%4et"
objUser.SetInfo

