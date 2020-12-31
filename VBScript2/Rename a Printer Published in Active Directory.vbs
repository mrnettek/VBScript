' Description: Uses the MoveHere method to rename a published printer in an OU.


Set objOU = GetObject("LDAP://ou=HR,dc=NA,dc=fabrikam,dc=com")

objOU.MoveHere _
    "LDAP://cn=Printer1,ou=HR,dc=NA,dc=fabrikam,dc=com", "cn=HRPrn1"

