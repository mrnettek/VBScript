' Description: Binds to the local user account on a computer named atl-win2k-01, and configures the account so that it never expires.


Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000

strComputer = "atl-win2k-01"
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer")

objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo

