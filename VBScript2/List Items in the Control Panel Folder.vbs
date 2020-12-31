' Description: Reports the path to the Windows Control Panel, and then lists the individual applications installed. For Windows NT 4.0 and Windows 98, this script requires Windows Script Host 5.1 and Internet Explorer 4.0 or later.


Const CONTROL_PANEL = &H3&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CONTROL_PANEL)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

