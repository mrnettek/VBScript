' Description: Uses the MoveHere method  to move an object to another domain. Note that there are a number of restrictions associated with performing this type of move operation. For details, see the Directory Services Platform SDK.


Set objOU = GetObject("LDAP://cn=Computers,dc=NA,dc=fabrikam,dc=com")

objOU.MoveHere "LDAP://cn=Computer01,cn=Users,dc=fabrikam,dc=com", _
    vbNullString

