' Description: Identifies whether a user account is enabled or disabled.


Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
If objUser.AccountDisabled = FALSE Then
      WScript.Echo "The account is enabled."
Else
      WScript.Echo "The account is disabled."
End If

