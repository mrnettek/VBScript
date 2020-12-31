' Description: Deletes an ADAM OU named Accounting.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/dc=fabrikam,dc=com")
objOU.Delete "organizationalUnit", "ou=Accounting"

