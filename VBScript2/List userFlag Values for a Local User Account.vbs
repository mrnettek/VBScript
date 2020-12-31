' Description: Accesses the userAccountControl to retrieve attribute values for the local user account kenmyer on a computer named atl-win2k-01. These attribute values include such things as account status (enabled or disabled), whether the user requires a password and, if so, whether or not that password will ever expire.


Const ADS_UF_SCRIPT = &H0001 
Const ADS_UF_ACCOUNTDISABLE = &H0002 
Const ADS_UF_HOMEDIR_REQUIRED = &H0008 
Const ADS_UF_LOCKOUT = &H0010 
Const ADS_UF_PASSWD_NOTREQD = &H0020 
Const ADS_UF_PASSWD_CANT_CHANGE = &H0040 
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H0080 
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000 
Const ADS_UF_SMARTCARD_REQUIRED = &H40000 
Const ADS_UF_PASSWORD_EXPIRED = &H800000 
 
Set usr = GetObject("WinNT://atl-win2k-01/kenmyer")
flag = usr.Get("UserFlags")
 
If flag AND ADS_UF_SCRIPT Then
    Wscript.Echo "Logon script will be executed."
Else
    Wscript.Echo "Logon script will not be executed."
End If
 
If flag AND ADS_UF_ACCOUNTDISABLE Then
    Wscript.Echo "Account is disabled."
Else
    Wscript.Echo "Account is not disabled."
End If
 
If flag AND ADS_UF_HOMEDIR_REQUIRED Then
    Wscript.Echo "Home directory required."
Else
    Wscript.Echo "Home directory not required."
End If
 
If flag AND ADS_UF_PASSWD_NOTREQD Then
    Wscript.Echo "Password not required."
Else
    Wscript.Echo "Password required."
End If
 
If flag AND ADS_PASSWORD_CANT_CHANGE Then
    Wscript.Echo "User cannot change password."
Else
    Wscript.Echo "User can change password."
End If
 
If flag AND ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED Then
    Wscript.Echo "Encrypted password allowed."
Else
    Wscript.Echo "Encrypted password not allowed."
End If
 
If flag AND ADS_UF_DONT_EXPIRE_PASSWD Then
    Wscript.Echo "Password does not expire."
Else
    Wscript.Echo "Password expires."
End If
 
If flag AND ADS_UF_SMARTCARD_REQUIRED Then
    Wscript.Echo "Smartcard required for logon."
Else
    Wscript.Echo "Smart card not required for logon."
End If
 
If flag AND ADS_UF_PASSWORD_EXPIRED Then
    Wscript.Echo "Password has expired."
Else
    Wscript.Echo "Password has not expired."
End If

