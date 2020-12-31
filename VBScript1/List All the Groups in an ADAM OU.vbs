' Description: Lists all the group accounts in an ADAM OU named Accounting.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
objOU.Filter = Array("group")

For Each objUser in objOU
    Wscript.Echo objUser.Name
Next

