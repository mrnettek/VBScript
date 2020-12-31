' Description: Modifies the account expiration date for an ADAM user named Kenmyer.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=kenmyer,ou=Accounting,dc=fabrikam,dc=com")

objUser.AccountExpirationDate = "03/30/2005"
objUser.SetInfo

