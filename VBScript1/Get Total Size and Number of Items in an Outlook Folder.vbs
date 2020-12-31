Const olFolderInbox = 6
Const olFolderSentMail = 5

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
   
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
Set colItems = objFolder.Items
Wscript.Echo "No. of items in Inbox: " & colItems.Count

For Each objItem in colItems
    intSize = intSize + objItem.Size
Next

Wscript.Echo "Size of Inbox: " & Int(intSize / 1024) & " KB"

intSize = 0

Set objFolder = objNamespace.GetDefaultFolder(olFolderSentMail)
Set colItems = objFolder.Items
Wscript.Echo "No. of items in Sent Mail folder: " & colItems.Count

For Each objItem in colItems
    intSize = intSize + objItem.Size
Next

Wscript.Echo "Size of Sent Mail folder: " & Int(intSize / 1024) & " KB"
  


