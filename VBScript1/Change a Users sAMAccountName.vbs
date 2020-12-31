Set objUser = GetObject("cn=Ken Myer, ou=Finance, dc=Fabrikam, dc=com")

objUser.sAMAccountName = "Ken.Myer"
objUser.userPrincipalName = "Ken.Myer"
objUser.SetInfo
  


