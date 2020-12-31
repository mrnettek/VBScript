On Error Resume Next

Set objUser = GetObject("LDAP://cn=MyerKen,ou=Finance,dc=fabrikam,dc=com")

dtmAccountExpiration = objUser.AccountExpirationDate 
 
If Err.Number = -2147467259 OR dtmAccountExpiration = #1/1/1970# Then
    WScript.Echo "This account has no expiration date."
Else
    WScript.Echo "Account expiration date: " & objUser.AccountExpirationDate
End If
  


