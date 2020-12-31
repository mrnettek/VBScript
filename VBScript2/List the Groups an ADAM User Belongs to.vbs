' Description: Lists all the groups that the ADAM user Kenmyer is a member of.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://localhost:389/cn=kenmyer,ou=Accounting,dc=fabrikam,dc=com")
arrMembersOf = objUser.GetEx("memberOf")

For Each strMemberOf in arrMembersOf
  WScript.Echo strMemberOf
Next

