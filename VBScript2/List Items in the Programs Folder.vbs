' Description: Reports the path to the Recycle Bin, and then lists any items found there. For Windows NT 4.0 and Windows 98, this script requires Windows Script Host 5.1 and Internet Explorer 4.0 or later.


Const PROGRAMS = &H2&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(PROGRAMS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

