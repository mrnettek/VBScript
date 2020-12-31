' Description: Lists all the members of an ADAM group named Accountants.


On Error Resume Next

Set objGroup = GetObject _
    ("LDAP://localhost:389/cn=Accountants,ou=Accounting,dc=fabrikam,dc=com")

For Each objUser in objGroup.Members
    Wscript.Echo objUser.Name
Next

