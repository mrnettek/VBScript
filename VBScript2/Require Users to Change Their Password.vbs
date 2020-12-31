' Description: Forces a user to change their password the next time they logon.


Set objUser = GetObject _
    ("LDAP://CN=myerken,OU=management,DC=Fabrikam,DC=com")

objUser.Put "pwdLastSet", 0
objUser.SetInfo

