Const PROGRAM_FILES = &H26&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(PROGRAM_FILES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path
  


