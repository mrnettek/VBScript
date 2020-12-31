' Description: Creates a new organizational unit within Active Directory.


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")

Set objOU = objDomain.Create("organizationalUnit", "ou=Management")
objOU.SetInfo

