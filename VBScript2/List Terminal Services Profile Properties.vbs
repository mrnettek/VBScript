' Description: Returns Terminal Services profile information for the MyerKen Active Directory user account.


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
WScript.echo "Terminal Services Profile Path : " & _
    objUser.TerminalServicesProfilePath 
WScript.echo "Terminal Services Home Directory: " & _
    objUser.TerminalServicesHomeDirectory
WScript.echo "Terminal Services Home Drive: " & _
    objUser.TerminalServicesHomeDrive
WScript.echo "Allow Logon: " & objUser.AllowLogon

