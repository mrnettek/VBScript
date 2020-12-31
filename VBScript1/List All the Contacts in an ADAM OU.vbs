' Description: Lists all the contacts in an ADAM OU named Accounting.


On Error ResumeNext

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
objOU.Filter = Array("contact")

For Each objUser in objOU
    Wscript.Echo objUser.Name
Next

