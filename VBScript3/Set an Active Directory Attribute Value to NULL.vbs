Const ADS_PROPERTY_CLEAR = 1 

Set objUser = GetObject _
   ("LDAP://cn=ken myer, ou=finance, dc=fabrikam, dc=com") 
 
objUser.PutEx ADS_PROPERTY_CLEAR, "telephoneNumber", 0
objUser.SetInfo
  


