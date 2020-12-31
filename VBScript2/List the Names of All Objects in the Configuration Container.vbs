' Description: Retrieves the names of the first-level objects in the Configuration container.


Set objConfiguration = GetObject _
    ("LDAP://cn=Configuration,dc=fabrikam,dc=com")
 
For Each objContainer in objConfiguration
    WScript.Echo objContainer.Name
Next

