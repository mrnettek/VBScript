Set objUser = GetObject("LDAP://CN=myerken,OU=management,DC=Fabrikam,DC=com")
Wscript.Echo "Password last changed: " & objUser.PasswordLastChanged
  


