Const olFolderSentMail  = 5

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")

Set objFolder = objNamespace.GetDefaultFolder(olFolderSentMail)

Set colItems = objFolder.Items

For i = colItems.Count to 1 Step - 1
    colItems(i).Delete
Next
  


