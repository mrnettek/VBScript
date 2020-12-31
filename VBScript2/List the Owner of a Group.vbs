' Description: Returns the owner of an Active Directory security group named Scientists.


Set objGroup = GetObject _
  ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
 
Set objNtSecurityDescriptor = objGroup.Get("ntSecurityDescriptor")
 
WScript.Echo "Owner Tab"
WScript.Echo "Current owner of this item: " & objNtSecurityDescriptor.Owner

