' Description: Deletes an organizational unit named HR from the domain fabrikam.com.


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")

objDomain.Delete "organizationalUnit", "ou=hr"

