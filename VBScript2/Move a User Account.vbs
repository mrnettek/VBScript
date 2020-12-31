' Description: Moves a user account from one OU to another.


Set objOU = GetObject("LDAP://ou=sales,dc=na,dc=fabrikam,dc=com")

objOU.MoveHere _
    "LDAP://cn=BarrAdam,OU=hr,dc=na,dc=fabrikam,dc=com", vbNullString

