' Description: Reports the path to the Fonts folder, and then lists any fonts found in that folder. For Windows NT 4.0 and Windows 98, this script requires Windows Script Host 5.1 and Internet Explorer 4.0 or later.


Const FONTS = &H14&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(FONTS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

