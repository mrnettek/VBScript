' Description: Returns Terminal Services environment properties for the MyerKen Active Directory user account.


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
Wscript.Echo "Connect Client Drives At Logon: " & _
    objUser.ConnectClientDrivesAtLogon
Wscript.Echo "Connect Client Printers At Logon: " & _
    objUser.ConnectClientPrintersAtLogon
Wscript.Echo "Default To Main Printer: " & objUser.DefaultToMainPrinter
Wscript.Echo "Terminal Services Initial Program: " & _
    objUser.TerminalServicesInitialProgram 
Wscript.Echo "Terminal Services Work Directory: " & _
    objUser.TerminalServicesWorkDirectory

