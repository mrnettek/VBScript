' Description: Creates an ADAM container named Users.


On Error Resume Next

Set objDomain = GetObject("LDAP://localhost:389/dc=fabrikam,dc=com")
Set objOU = objDomain.Create("container", "cn=Users")
objOU.SetInfo

