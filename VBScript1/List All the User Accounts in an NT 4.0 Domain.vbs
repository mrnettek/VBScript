' Description: Returns a list of all the user accounts in a Windows NT 4.0 domain named Fabrikam.


Set objDomain = GetObject("WinNT://fabrikam,domain")
objDomain.Filter = Array("User")

For Each objUser In objDomain
    Wscript.Echo objUser.Name 
Next

