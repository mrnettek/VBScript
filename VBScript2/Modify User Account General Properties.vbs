' Description: Configures general attributes for a user account.


Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
objUser.Put "userPrincipalName", "MyerKen@fabrikam.com"
objUser.Put "sAMAccountName", "MyerKen01"
objUser.Put "userWorkstations", "wks1,wks2,wks3"
objUser.SetInfo

