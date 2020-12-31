' Description: Returns a list of Shell object verbs (context menu items) for the Recycle Bin.


Const RECYCLE_BIN = &Ha&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(RECYCLE_BIN) 
Set objFolderItem = objFolder.Self      
Set colVerbs = objFolderItem.Verbs

For i = 0 to colVerbs.Count - 1
    Wscript.Echo colVerbs.Item(i)
Next

