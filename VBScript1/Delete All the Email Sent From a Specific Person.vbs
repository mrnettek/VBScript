Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[From] = 'Bill Gates'")

For i = colFilteredItems.Count to 1 Step -1
    colFilteredItems(i).Delete
Next
  


