Set objShell = CreateObject("Wscript.Shell")
For i = 1 to 1023
   strMessage = strMessage & "."
Next
strMessage = strMessage & "X"

MsgBox strMessage
Wscript.Echo strMessage
  


