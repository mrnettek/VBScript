Set objOU = GetObject("LDAP://ou=Finance,dc=fabrikam,dc=com")

If objOU.gpOptions = 1 Then
    Wscript.Echo "Block policy inheritance is enabled."
Else
    Wscript.Echo "Block policy inheritance is not enabled."
End If
  


