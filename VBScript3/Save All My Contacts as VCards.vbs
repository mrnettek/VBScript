On Error Resume Next

Const olFolderContacts = 10
Const olVCard = 6

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")

Set colContacts = objNamespace.GetDefaultFolder(olFolderContacts).Items

For Each objContact In colContacts
    strName = objContact.FirstName & objContact.LastName
    strPath = "C:\Test\" & strName & ".vcf"
    objContact.SaveAs strpath, olVCard
Next
  


