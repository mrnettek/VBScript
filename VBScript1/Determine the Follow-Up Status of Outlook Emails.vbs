Const olFolderInbox = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)

Set colItems = objFolder.Items

For Each objItem in colItems
    If objItem.FlagStatus = 1 Then
        Wscript.Echo "Follow-up complete"
        Wscript.Echo objItem.Subject 
        Wscript.Echo
    End If

    If objItem.FlagStatus = 2 Then
        Wscript.Echo "Marked for follow-up"
        Wscript.Echo objItem.Subject 
        Wscript.Echo
    End If
Next
  


