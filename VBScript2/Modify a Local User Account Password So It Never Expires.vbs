' Description: Binds to a local user account on a computer named atl-win2k-01, and configures the account password so  it never expires.


Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
strDomainOrWorkgroup = "Fabrikam"
strComputer = "atl-win2k-01"
strUser = "KenMeyer"
 
Set objUser = GetObject("WinNT://" & strDomainOrWorkgroup & "/" & _
    strComputer & "/" & strUser & ",User")
 
objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo

