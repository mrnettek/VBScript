' Description: Configures a new password for a user.


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=management,dc=fabrikam,dc=com")

objUser.SetPassword "i5A2sj*!"

