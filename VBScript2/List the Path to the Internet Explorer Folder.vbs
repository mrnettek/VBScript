' Description: Reports the path to the Internet Explorer special folder. For Windows NT 4.0 and Windows 98, this script requires Windows Script Host 5.1 and Internet Explorer 4.0 or later.


Const INTERNET_EXPLORER = &H1&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(INTERNET_EXPLORER)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

