' Description: Binds to the local Administrator account on the computer atl-ws-01, and changes the password for the account to 09iuy%4e.


strComputer = "atl-ws-01"
Set objUser = GetObject("WinNT://" & strComputer & "/Administrator, user")

objUser.SetPassword "09iuy%4e"
objUser.SetInfo

