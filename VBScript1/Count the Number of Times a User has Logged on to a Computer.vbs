Set objUser = GetObject _
    ("LDAP://atl-dc-01/cn=ken myer, ou=Finance, dc=fabrikam, dc=com")
Wscript.Echo objUser.LogonCount
  


