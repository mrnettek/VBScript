' Description: Uses the Shell object to print all the files in the C:\Logs folder.


TargetFolder = "C:\Logs" 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(TargetFolder) 
Set colItems = objFolder.Items

For i = 0 to colItems.Count - 1
    colItems.Item(i).InvokeVerbEx("Print")
Next

