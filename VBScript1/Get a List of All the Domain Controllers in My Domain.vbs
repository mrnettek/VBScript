Set objOU = GetObject(“LDAP://ou=Domain Controllers, dc=fabrikam, dc=com”)
objOU.Filter = Array(“Computer”)
For Each objComputer in objOU
    Wscript.Echo objComputer.CN
Next
  


