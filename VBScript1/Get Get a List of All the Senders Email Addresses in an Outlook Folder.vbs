Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colItems = objFolder.Items

For Each objItem in colItems
    Wscript.Echo objItem.SenderEmailAddress
Next
  


