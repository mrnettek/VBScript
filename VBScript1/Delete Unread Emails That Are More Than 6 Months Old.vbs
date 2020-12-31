Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[UnRead] = True")

For i = colFilteredItems.Count to 1 Step - 1
    If DateDiff("m", colFilteredItems(i).ReceivedTime, Now) > 6 Then
        colFilteredItems(i).Delete
    End If
Next
  


