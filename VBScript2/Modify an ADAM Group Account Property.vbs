' Description: Modifies the description attribute for an ADAM group named Accountants.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=Accountants,ou=Accounting,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice group account."  
objUser.SetInfo

