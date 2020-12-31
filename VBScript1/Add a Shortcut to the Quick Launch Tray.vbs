Set objShell = CreateObject("WScript.Shell")
Set colEnvironmentVariables = objShell.Environment("Volatile")

strFolder = colEnvironmentVariables.Item("APPDATA") & _
 "\Microsoft\Internet Explorer\Quick Launch"

Set objShortCut = objShell.CreateShortcut(strFolder & _
 "\Notepad.lnk")
objShortCut.TargetPath = "Notepad.exe"
objShortCut.Description = "Open Notepad"
objShortCut.HotKey = "Ctrl+Shift+N"
objShortCut.Save
  


