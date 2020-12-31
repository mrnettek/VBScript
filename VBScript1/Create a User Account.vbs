' Description: Creates a user account in Active Directory. This script only creates the account, it does not enable it.


Set objOU = GetObject("LDAP://OU=management,dc=fabrikam,dc=com")

Set objUser = objOU.Create("User", "cn=MyerKen")
objUser.Put "sAMAccountName", "myerken"
objUser.SetInfo

