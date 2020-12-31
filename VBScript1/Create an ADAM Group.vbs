' Description: Creates a new ADAM group named Accountants.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("group", "cn=Accountants")
objGroup.SetInfo

