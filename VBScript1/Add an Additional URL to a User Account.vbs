' Description: Adds an additional URL to a user account. Demonstrates how to append a new value to a multi-valued attribute.


Const ADS_PROPERTY_APPEND = 3 
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_APPEND, _
    "url", Array("http://www.fabrikam.com/policy")
objUser.SetInfo

