On Error Resume Next

Set objUser = GetObject("LDAP://cn=Ken Myer,ou=Finance,dc=fabrikam,dc=com")

i = 0

For Each strGroup in objUser.memberOf
    Set objGroup = GetObject("LDAP://" &  strGroup)
    If objGroup.CN = "Finance Users" Then
        i = i + 1
    End If
    If objGroup.CN = "Fabrikam Managers" Then
        i = i + 1
    End If
Next

If i = 2 Then
    Set objGroup = GetObject("LDAP://cn=Finance Managers,ou=Finance,dc=fabrikam,dc=com")
    objGroup.Add(objUser.ADsPath)
End If
  


