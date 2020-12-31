' Description: Configures the MyerKen Active Directory user account to expire on March 30, 2005.


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

objUser.AccountExpirationDate = "03/30/2005"
objUser.SetInfo

