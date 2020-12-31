On Error Resume Next

Set objUser = GetObject("LDAP://cn=Ken Myer, ou=Finance, dc=fabrikam, dc=com")

If IsEmpty(objUser.homeDirectory) or IsNull(objUser.homeDirectory) Then
    strUser = objUser.sAMAccountName
    strHomeDirectory = "\\atl-dc-01\users\" & strUser
    objUser.homeDirectory = strHomeDirectory
    objUser.homeDrive = "Z:"
    objUser.SetInfo
End If
  


