Set objUser = GetObject _
    ("LDAP://CN=Ken Myer,OU=Finance,DC=fabrikam,DC=com")
Wscript.Echo objUser.department
  


