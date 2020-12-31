strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colFiles = objWMIService.ExecQuery("Select * From Win32_ShortcutFile")
For Each objFile in colFiles
    Wscript.Echo "Name: " & objFile.FileName
    Wscript.Echo "Shortcut target: " & objFile.Target
    Wscript.Echo "File name: " & objFile.Description
Next
  


