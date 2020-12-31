' Description: Changes the server name portion of the user profile path to \\fabrikam for the MyerKen Active Directory user account.


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
strCurrentProfilePath = objUser.Get("profilePath")
intStringLen = Len(strCurrentProfilePath)
intStringRemains = intStringLen - 11
strRemains = Mid(strCurrentProfilePath, 12, intStringRemains)
strNewProfilePath = "\\fabrikam" & strRemains
objUser.Put "profilePath", strNewProfilePath
objUser.SetInfo

