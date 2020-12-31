Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[SenderEmailAddress] = 'kenmyer@fabrikam.com'")

For Each objMessage In colFilteredItems
    Wscript.Echo objMessage.Subject
Next
  


