' Description: Modifies the description attribute for an ADAM OU named Accounting.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice organizational unit."  
objUser.SetInfo

