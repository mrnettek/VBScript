Set objGroup = GetObject("LDAP://CN=Managers,OU=Finance,DC=fabrikam,DC=com")

For Each objUser in objGroup.Members
    Wscript.Echo "Name: " & objUser.DisplayName
    Wscript.Echo "Department: " & objUser.department
    Wscript.Echo "Street address: " & objUser.streetAddress
    Wscript.Echo "Title: " & objUser.title
    Wscript.Echo "Description: " & objUser.description
    Wscript.Echo
Next
  


