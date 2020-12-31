Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &H8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &H80000000
 
Set objGroup = GetObject _
    ("LDAP://cn=Managers,ou=Finance,dc=fabrikam,dc=com") 
 
objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP OR ADS_GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo
  


