' Description: Removes user Jackson from the group Office-Users.


Const ADS_PROPERTY_DELETE = 4 
 
Set objGroup = GetObject _
   ("LDAP://cn=Office-Users,cn=Users,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_DELETE, _
    "member",Array("cn=Jackson,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo

