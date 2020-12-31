' Description: Binds to a local user account (kenmyer) on a computer named atl-win2k-01, and configures the account to expire on March 1, 2005.


strComputer = "atl-win2k-01"
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer")

objUser.AccountExpirationDate = #03/01/2005# 
objUser.SetInfo

