' Description: Deletes the user account MyerKen from the HR organizational unit in a domain named fabrikam.com.


Set objOU = GetObject("LDAP://ou=hr,dc=fabrikam,dc=com")

objOU.Delete "user", "cn=MyerKen"

