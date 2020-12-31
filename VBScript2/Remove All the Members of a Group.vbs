' Description: Removes all the members of an Active Directory group named Office-Users.


Const ADS_PROPERTY_CLEAR = 1 
 
Set objGroup = GetObject _
    ("LDAP://cn=Office-Users,cn=Users,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
objGroup.SetInfo

