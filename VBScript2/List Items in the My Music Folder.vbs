' Description: Reports the path to the My Music folder, and then lists any items found in that folder.


Const MY_MUSIC = &Hd&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_MUSIC)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

