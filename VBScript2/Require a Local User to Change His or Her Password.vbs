' Description: Binds to a local user account on the computer named atl-win2k-01, and configures the account so that the user (kenmyer) must change his password the next time he logs on.


strComputer = "atl-win2k-01"
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer")

objUser.Put "PasswordExpired", 1
objUser.SetInfo

