' Description: Configures Terminal Services session attributes for the MyerKen Active Directory user account.


Const Enabled = 1
Const Disabled = 0

Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.MaxDisconnectionTime = 2880
objUser.MaxConnectionTime = 1440
objUser.MaxIdleTime = 180
objUser.BrokenConnectionAction = Enabled
objUser.ReconnectionAction = Enabled
objUser.SetInfo

