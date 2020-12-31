' Description: Retrieves Terminal Services session properties for the MyerKen Active Directory user account.


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
WScript.echo "Maximum Disconnection Time : " & objUser.MaxDisconnectionTime 
WScript.echo "Maximum Connection Time: " & objUser.MaxConnectionTime
WScript.echo "Maximum Idle Time: " & objUser.MaxIdleTime
WScript.echo "Broken Connection Action: " & objUser.BrokenConnectionAction 
WScript.echo "Reconnection Action : " & objUser.ReconnectionAction

