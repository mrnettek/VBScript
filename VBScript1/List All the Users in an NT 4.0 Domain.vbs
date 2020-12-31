Set objDomain = GetObject("WinNT://fabrikam")
objDomain.Filter = Array("User")

For Each objUser in objDomain
    Wscript.Echo "User name: " & objUser.Name 
    Wscript.Echo "Description: " & objUser.Description 
    Wscript.Echo "Logon script path: " & objUser.LoginScript 
    Wscript.Echo 
Next
  


