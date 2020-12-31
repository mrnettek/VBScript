' Description: Retrieves user account attributes found on the Account page of the user account object in Active Directory Users and Computers.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
WScript.Echo "User Principal Name: " & objUser.userPrincipalName
WScript.Echo "SAM Account Name: " & objUser.sAMAccountName
WScript.Echo "User Workstations: " & objUser.userWorkstations

Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")
WScript.Echo "Domain controller: " & objDomain.dc

