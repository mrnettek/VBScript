' Description: Returns information found on the General Properties page in Active Directory Users and Computers for an OU named Sales.


On Error Resume Next

Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
For Each strValue in objContainer.description
  WScript.Echo "Description: " & strValue
Next
 
Wscript.Echo "Street Address: " & strStreetAddress
Wscript.Echo "Locality: " & 
Wscript.Echo "State/porvince: " & objContainer.st
Wscript.Echo "Postal Code: " & objContainer.postalCode
Wscript.Echo "Country: " & objContainer.c

