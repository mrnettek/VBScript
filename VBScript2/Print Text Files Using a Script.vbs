TargetFolder = "C:\Logs" 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(TargetFolder) 
Set colItems = objFolder.Items
For Each objItem in colItems
    objItem.InvokeVerbEx("Print")
Next
  


