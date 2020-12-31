' Description: Modifies the description attribute for an ADAM inetOrgPerson account named Syedabbas.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=syedabbas,ou=Accounting,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice inetOrgPerson account."  
objUser.SetInfo

