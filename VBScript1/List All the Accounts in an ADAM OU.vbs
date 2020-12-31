' Description: Lists all the accounts (of any type) in an ADAM OU named Accounting.


On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")

For Each objUser in objOU
    Wscript.Echo objUser.Name
Next

