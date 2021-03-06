' Description: Appends a new phone number to the otherHomePhone attribute of an Active Directory user account. This operation adds the phone number to the attribute without deleting any existing phone numbers.


Const ADS_PROPERTY_APPEND = 3 
 
Set objUser = GetObject _
   ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 

objUser.PutEx ADS_PROPERTY_APPEND, "otherHomePhone", Array("(425) 555-0116")
objUser.SetInfo

