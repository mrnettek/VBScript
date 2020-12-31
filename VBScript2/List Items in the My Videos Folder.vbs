' Description: Reports the path to the My Videos folder, and then lists any items found in that folder.


Const MY_VIDEOS = &He&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_VIDEOS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

