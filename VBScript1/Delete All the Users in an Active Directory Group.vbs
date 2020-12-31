Const ADS_PROPERTY_CLEAR = 1 
 
Set objGroup = GetObject("LDAP://cn=Finance Users,ou=Finance,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
objGroup.SetInfo
  


