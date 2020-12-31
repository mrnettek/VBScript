Set objUser = GetObject _
    ("LDAP://cn=Ken Myer,ou=Finance,dc=fabrikam,dc=com")
 
Set objGroup = GetObject _
    ("LDAP://cn=Finance Managers,ou=Finance,dc=fabrikam,dc=com")

objGroup.GetInfoEx Array("primaryGroupToken"), 0

objUser.primaryGroupID = objGroup.primaryGroupToken
objUser.SetInfo
  


