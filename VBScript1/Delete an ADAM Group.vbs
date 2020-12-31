' Description: Deletes an ADAM group named Accountants.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
objOU.Delete "group", "cn=Accountants"

