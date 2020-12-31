On Error Resume Next

Const olFolderContacts = 10

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")

Set colContacts = objNamespace.GetDefaultFolder(olFolderContacts).Items

For Each objContact In colContacts
    If Month(objContact.Birthday) = Month(Date) Then
        Wscript.Echo objContact.FullName, objContact.Birthday
    End If
Next
  


