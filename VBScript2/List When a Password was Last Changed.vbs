' Description: Identifies the last time a user password was changed.


Set objUser = GetObject _
    ("LDAP://CN=myerken,OU=management,DC=Fabrikam,DC=com")

dtmValue = objUser.PasswordLastChanged
WScript.Echo "Password last changed: " & dtmValue

