' Description: Deletes an ADAM user account named Carolphilips.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
objOU.Delete "user", "cn=carolphilips"

