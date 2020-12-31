' Description: Retrieves the ADsPath, Class, GUID, Name, Parent, and Schema properties  for a domain.


Set objDomain = GetObject("LDAP://dc=NA,dc=fabrikam,dc=com")

WScript.Echo "Ads Path:" & objDomain.ADsPath
WScript.Echo "Class:" & objDomain.Class
WScript.Echo "GUID:" & objDomain.GUID
WScript.Echo "Name:" & objDomain.Name
WScript.Echo "Parent:" & objDomain.Parent
WScript.Echo "Schema:" & objDomain.Schema

