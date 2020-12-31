Const olFolderInbox = 6
Const olTxt = 0

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colMailItems =  objFolder.Items

Set objItem = colMailItems.GetLast()
objItem.SaveAs "C:\Scripts\MailMessage.txt", olTxt
  


