Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")

strPath = objShell.CurrentDirectory
strDrive = objFSO.GetDriveName(strPath)

Wscript.Echo strDrive
  


