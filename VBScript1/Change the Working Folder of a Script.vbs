Set objShell = CreateObject("Wscript.Shell")
Wscript.Echo objShell.CurrentDirectory

objShell.CurrentDirectory = "C:\Windows"
Wscript.Echo objShell.CurrentDirectory
  


