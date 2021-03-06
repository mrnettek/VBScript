' Description: Configures the EnableRemoteControl attribute for the MyerKen Active Directory user account. Other constant values are included in the script as a reference.


Const Disable = 0
Const EnableInputNotify = 1
Const EnableInputNoNotify = 2 
Const EnableNoInputNotify = 3
Const EnableNoInputNoNotify = 4
 
Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objUser.EnableRemoteControl = EnableNoInputNoNotify
 
objUser.SetInfo

