Const olFolderCalendar = 9

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)

Set colItems = objFolder.Items

strFilter = "[Organizer] = 'Ken Myer'"
Set colFilteredItems = colItems.Restrict(strFilter)

For Each objItem In colFilteredItems
    If objItem.Start > Now Then
        Wscript.Echo "Meeting name: " & objItem.Subject
        Wscript.Echo "Meeting date: " & objItem.Start
        Wscript.Echo "Duration: " & objItem.Duration & " minutes"
        Wscript.Echo "Location: " & objItem.Location
        Wscript.Echo
    End If
Next
  


