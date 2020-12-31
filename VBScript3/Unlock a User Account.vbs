' Description: Unlocks the MyerKen Active Directory user account.


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

objUser.IsAccountLocked = False
objUser.SetInfo

