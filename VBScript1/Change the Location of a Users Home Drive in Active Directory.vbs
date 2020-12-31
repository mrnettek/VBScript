Set objUser = GetObject("LDAP://cn=Ken Myer,ou=finance,dc=fabrikam,dc=com")
 
objUser.Put "homeDirectory", "\\atl-fs-01\users\kenmyer"
objUser.Put "homeDrive", "X:"

objUser.SetInfo
  


