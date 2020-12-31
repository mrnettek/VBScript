' Description: Modifies the description attribute for an ADAM user account named Kenmyer.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=kenmyer,ou=Accounting,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice user account."  
objUser.SetInfo

