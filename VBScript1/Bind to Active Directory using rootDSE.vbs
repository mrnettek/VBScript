' Description: Demonstration script that uses rootDSE to bind to various objects in the local Active Directory domain.


Set objRootDSE = GetObject("LDAP://rootDSE")
 
strSchema = "LDAP://" & objRootDSE.Get("schemaNamingContext")
WScript.echo "ADsPath to schema: " & strSchema
Set objSchema = GetObject(strSchema)
WScript.Echo "Schema Object:"
WScript.Echo "Name: " & objSchema.Name
WScript.Echo "Class: " & objSchema.Class & VbCrLf
 
strConfiguration = "LDAP://" & objRootDSE.Get("configurationNamingContext")
WScript.Echo "ADsPath to configuration container: " & strConfiguration
Set objConfiguration = GetObject(strConfiguration)
WScript.Echo "Configuration Object:"
WScript.Echo "Name: " & objConfiguration.Name
WScript.Echo "Class: " & objConfiguration.Class & VbCrLf
 
strDomain = "LDAP://" & objRootDSE.Get("defaultNamingContext")
WScript.Echo "ADsPath to current domain container: " & strDomain
Set objDomain = GetObject(strDomain)
WScript.Echo "Current Domain Object:"
WScript.Echo "Name: " & objDomain.Name
WScript.Echo "Class: " & objDomain.Class & VbCrLf
 
strRootDomain = "LDAP://" & objRootDSE.Get("rootDomainNamingContext")
WScript.Echo "ADsPath to root domain container: " & strDomain
Set objRootDomain = GetObject(strRootDomain)
WScript.Echo "Current Domain Object:"
WScript.Echo "Name: " & objRootDomain.Name
WScript.Echo "Class: " & objRootDomain.Class & VbCrLf

