' Description: Retrieves Terminal Services Remote Control attribute values for the MyerKen user account.


Const Disable = 0
Const EnableInputNotify = 1
Const EnableInputNoNotify = 2 
Const EnableNoInputNotify = 3
Const EnableNoInputNoNotify = 4
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
intEnableRemoteControl  = objUser.EnableRemoteControl  
 
Select Case intEnableRemoteControl
    Case Disable  WScript.Echo "Remote control disabled"
    Case EnableInputNotify 
        WScript.Echo "Remote control enabled"
        WScript.Echo "User permission required"
        WScript.Echo "Interact with the session"
    Case EnableInputNoNotify
        WScript.Echo "Remote control enabled"
        WScript.Echo "User permission not required"
        WScript.Echo "Interact with the session"
    Case EnableNoInputNotify
        WScript.Echo "Remote control enabled"
        WScript.Echo "User permission required"
        WScript.Echo "View the session"
    Case EnableNoInputNoNotify
        WScript.Echo "Remote control enabled"
        WScript.Echo "User permission not required"
        WScript.Echo "View the session"
End Select

