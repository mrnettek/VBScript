If Wscript.Arguments.Named("s") = "" Or Wscript.Arguments.Named("c") = "" Then
    Wscript.Echo "You must specify both the service and the computer using syntax like this:"
    Wscript.Echo
    Wscript.Echo "myscript.vbs /s:alerter /c:atl-ws-01"
    Wscript.Quit
End If

Wscript.Echo "Service: " & Wscript.Arguments.Named("s")
Wscript.Echo "Computer: " & Wscript.Arguments.Named("c")
  


