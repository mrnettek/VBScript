Set objUser = GetObject("LDAP://cn=Jack Richins,ou=canada,dc=fabrikam,dc=com")

If objUser.Department = "IT" Then
    Set objGroup = GetObject _
        ("LDAP://cn=IT Staff,ou=support,dc=fabrikam,dc=com")
    objGroup.Add(objUser.ADsPath)
End If
  


