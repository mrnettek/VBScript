' Description: Adds the user Kenmyer to the ADAM Administrators group.


On Error Resume Next

Set objGroup = GetObject _
    ("LDAP://localhost:389/CN=Administrators,CN=Roles,dc=fabrikam,dc=com")
Set objUser = GetObject _
    ("LDAP://localhost:389/cn=kenmyer,ou=Accounting,dc=fabrikam,dc=com")
objGroup.Add objUser.AdsPath

