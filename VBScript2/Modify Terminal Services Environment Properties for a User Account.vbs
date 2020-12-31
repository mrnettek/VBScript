' Description: Configures the Terminal Services environment properties for the MyerKen Active Directory user account.


Const Enabled = 1
Const Disabled = 0

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.ConnectClientDrivesAtLogon = Enabled
objUser.ConnectClientPrintersAtLogon = Enabled
objUser.DefaultToMainPrinter = Enabled
objUser.TerminalServicesInitialProgram = "cmd"
objUser.TerminalServicesWorkDirectory = "c:\temp"
objUser.SetInfo

