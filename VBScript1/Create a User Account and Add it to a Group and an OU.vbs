' Description: Demonstration script that: 1) creates a new Active Directory organizational unit; 2) creates a new user account and new security group; and, 3) adds the new user as a member of that security group.


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")
Set objOU = objDomain.Create("organizationalUnit", "ou=Management")
objOU.SetInfo
 
Set objOU = GetObject("LDAP://OU=Management,dc=fabrikam,dc=com")
Set objUser = objOU.Create("User", "cn= AckermanPilar")
objUser.Put "sAMAccountName", "AckermanPila"
objUser.SetInfo
 
Set objOU = GetObject("LDAP://OU=Management,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=atl-users")
objGroup.Put "sAMAccountName", "atl-users"
objGroup.SetInfo
 
objGroup.Add objUser.ADSPath

