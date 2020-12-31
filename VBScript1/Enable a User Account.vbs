' Description: Enables a user account.


Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")

objUser.AccountDisabled = FALSE
objUser.SetInfo

