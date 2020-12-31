Set objUser = GetObject("LDAP://CN=myerken,OU=Finance,DC=Fabrikam,DC=com")

objUser.pwdLastSet = 0
objUser.SetInfo
  


