Set objShell = CreateObject("Wscript.Shell")

intMessage = Msgbox("Would you like to apply for access to this resource?", _
    vbYesNo, "Access Denied")

If intMessage = vbYes Then
    objShell.Run("http://www.microsoft.com")
Else
    Wscript.Quit
End If
  


