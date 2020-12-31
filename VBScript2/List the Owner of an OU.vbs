' Description: Returns the owner of the Sales OU in Active Directory.


Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
Set objNtSecurityDescriptor = objContainer.Get("ntSecurityDescriptor")
 
WScript.Echo "Owner Tab"
WScript.Echo "Current owner of this item: " & objNtSecurityDescriptor.Owner

