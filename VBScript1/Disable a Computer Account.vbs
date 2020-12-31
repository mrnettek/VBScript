Set objComputer = GetObject("LDAP://cn=atl-ws-01,cn=computers,dc=fabrikam,dc=com")

objComputer.AccountDisabled = True
objComputer.SetInfo
  


