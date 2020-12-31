Const ADS_UF_PASSWD_NOTREQD = &H0020 
 
Set colUsers = GetObject("WinNT://atl-fs-01")
colUsers.Filter = Array("User")

For Each objUser in colUsers
   If objUser.UserFlags AND ADS_UF_PASSWD_NOTREQD Then
        Wscript.Echo objUser.Name & ": Password not required."
    Else
        Wscript.Echo objUser.Name & ": Password required."
    End If
Next
  


