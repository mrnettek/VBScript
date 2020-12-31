' Description: Reports the path to the Printers and Faxes folder, and then lists any items found in that folder. For Windows NT 4.0 and Windows 98, this script requires Windows Script Host 5.1 and Internet Explorer 4.0 or later.


Const PRINTERS_AND_FAXES = &H4&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(PRINTERS_AND_FAXES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

