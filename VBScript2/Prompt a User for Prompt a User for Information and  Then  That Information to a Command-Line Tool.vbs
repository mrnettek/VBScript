Set objShell = CreateObject("Wscript.Shell")

strName = InputBox("Please enter the user name:")

If strName = "" Then
    Wscript.Quit
End If

strCommand = "%comspec% /k dsquery user -name " & Chr(34) & strName & chr(34)
strCommand = strCommand & " | dsget user -tel"

objShell.Run strCommand
  


