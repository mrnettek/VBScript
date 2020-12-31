Set objShell = CreateObject("Wscript.Shell")

intValue = InputBox("Please enter a number:")
strCommandLine = "output.vbs " & intValue
objShell.Run(strCommandLine)
  


