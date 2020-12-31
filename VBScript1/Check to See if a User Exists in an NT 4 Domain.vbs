Set objDomain = GetObject("WinNT://fabrikam")
objDomain.Filter = Array("User")

blnFound = FALSE

For Each objUser in objDomain
    If objUser.Name = "kenmyer" Then
        blnFound = TRUE
        Exit For
    End If 
Next

If blnFound = TRUE Then
    Wscript.Echo "The user account exists in the domain."
Else
    Wscript.Echo "The user account does not exist in the domain."
End If
  


