' Description: Uses the MoveHere method to move a user account to another domain. Note that there are a number of restrictions associated with performing this type of move operation.


Set objOU = GetObject("LDAP://ou=management,dc=na,dc=fabrikam,dc=com")

objOU.MoveHere _
    "LDAP://cn=AckermanPilar,OU=management,dc=fabrikam,dc=com", vbNullString

