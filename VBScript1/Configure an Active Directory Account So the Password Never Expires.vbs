Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
Set objUser = GetObject("LDAP://cn=Ken Myer, ou=Finance, dc=fabrikam, dc=com")

intUserAccountControl = objUser.Get("userAccountControl")
 
If Not objUser.userAccountControl AND ADS_UF_DONT_EXPIRE_PASSWD Then
    objUser.Put "userAccountControl", _
        objUser.userAccountControl XOR ADS_UF_DONT_EXPIRE_PASSWD
    objUser.SetInfo
End If
  


