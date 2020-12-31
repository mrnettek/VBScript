Const DESKTOP = &H10&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(DESKTOP)
Set objFolderItem = objFolder.Self

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next
  


