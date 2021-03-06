' Description: Configures the MyerKen user account so that the user must use a smartcard in order to logon to Active Directory.


Const ADS_UF_SMARTCARD_REQUIRED = &h40000 

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
intUAC = objUser.Get("userAccountControl")
 
If (intUAC AND ADS_UF_SMARTCARD_REQUIRED) = 0 Then
    objUser.Put "userAccountControl", intUAC XOR ADS_UF_SMARTCARD_REQUIRED
    objUser.SetInfo
End If

