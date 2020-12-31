' Description: Retrieves user account attributes found on the General Properties page of the user account object in Active Directory Users and Computers.


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")


WScript.Echo "First Name: " & objUser.givenName
WScript.Echo "Initials: " & objUser.initials
WScript.Echo "Last Name: " & objUser.sn
WScript.Echo "Display Name: " & objUser.displayName
WScript.Echo "Office: " & _
    objUser.physicalDeliveryOfficeName
WScript.Echo "Telephone Number: " & objUser.telephoneNumber
WScript.Echo "Email: " & objUser.mail
WScript.Echo "Home Page: " & 
 
For Each strValue in objUser.description
    WScript.Echo "Description: " & strValue
Next

For Each strValue in objUser.otherTelephone
    WScript.Echo "Other Telephone: " & strValue
Next

For Each strValue in objUser.url
    WScript.Echo "URL: " & strValue
Next

