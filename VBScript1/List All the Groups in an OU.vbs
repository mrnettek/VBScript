Set objOU = GetObject("LDAP://ou=Finance,dc=fabrikam,dc=com")
objOU.Filter = Array("Group")

For Each objGroup in objOU
    Wscript.Echo objGroup.Name
Next
  


