' Description: Modifies the description attribute for an ADAM contact named Carolphilips.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=carolphilips,ou=Accounting,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice contact account."  
objUser.SetInfo

