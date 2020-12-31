' Description: Creates a new ADAM OU named Management.


On Error Resume Next

Set objDomain = GetObject("LDAP://localhost:389/dc=fabrikam,dc=com")
Set objOU = objDomain.Create("organizationalUnit", "ou=Management")
objOU.SetInfo

