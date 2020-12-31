Const olFolderContacts  = 10

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderContacts)

Set objList = objFolder.Items("Approved Vendors")

For i = 1 to objList.MemberCount
    Set objMember = objList.GetMember(i)
    Wscript.Echo objMember.Name & ", " & objMember.Address 
Next
  


