Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objInbox = objNamespace.GetDefaultFolder(olFolderInbox)

strFolderName = objInbox.Parent

Set objMailbox = objNamespace.Folders(strFolderName)

Set objFolder = objMailbox.Folders("Europe")

Set colItems = objFolder.Items

For Each objItem in colItems
    Wscript.Echo objItem.Subject
Next
  


