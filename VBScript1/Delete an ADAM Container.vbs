' Description: Deletes an ADAM container named Users.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
objOU.Delete "container", "cn=Users"

