Set objOU = GetObject("LDAP://ou=Accounting,dc=fabrikam,dc=com")
objOU.Filter = Array("user")

For Each objUser in objOU
    objUser.pwdLastSet = 0
    objUser.SetInfo
Next
  


