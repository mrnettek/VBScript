' Description: Creates an ADAM contact named Jonathanhaas.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
Set objUser = objOU.Create("contact", "cn=jonathanhaas")
objUser.Put "displayName", "Jonathan Haas"  
objUser.SetInfo

