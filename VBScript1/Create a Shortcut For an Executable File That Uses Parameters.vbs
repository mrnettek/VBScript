Set objShell = CreateObject("Wscript.Shell")
strDesktop = objShell.SpecialFolders("Desktop")

Set objShortcut = objShell.CreateShortcut(strDesktop & "\Test.lnk")
objShortcut.TargetPath = "Notepad.exe"
objShortcut.Arguments = "C:\Scripts\Test.txt"

objShortcut.Description = "Starts Notepad with a file already loaded."
objShortcut.WorkingDirectory = strDesktop

objShortcut.Save
  


