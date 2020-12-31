' Description: Retrieves user account attributes found on the Address page of the user account object in Active Directory Users and Computers.


On Error Resume Next
 
Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
WScript.Echo "Street Address: " & objUser.streetAddress
WScript.Echo "Post Office Box: " & objUser.postOfficeBox
WScript.Echo "Locality: " & objUser.l
WScript.Echo "Street: " & objUser.st
WScript.Echo "Postal Code: " & objUser.postalCode
WScript.Echo "Country: " & objUser.c

