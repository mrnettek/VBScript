' Description: Configures Terminal Services profile attributes for the MyerKen Active Directory user account.


Const Enabled = 1
Const Disabled = 0

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.TerminalServicesProfilePath = ""
objUser.TerminalServicesHomeDirectory = ""
objUser.TerminalServicesHomeDrive = ""
objUser.AllowLogon = Enabled
objUser.SetInfo

