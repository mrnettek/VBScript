' Description: Modifies the description attribute for an ADAM container named Users.


On Error Resume Next

Set objUser = GetObject("LDAP://localhost:389/cn=users,dc=fabrikam,dc=com")

objUser.Put "description", "This is a practice container."  
objUser.SetInfo

