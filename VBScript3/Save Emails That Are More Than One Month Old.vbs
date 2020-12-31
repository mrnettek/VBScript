On Error Resume Next

Const olFolderSentMail = 5
Const olMSG = 3

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderSentMail)

dtmTargetDate = Date - 30

Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[CreationTime] <'" & dtmTargetDate & "'")

For Each objMessage In colFilteredItems
    strName = objMessage.Subject
    strName = Replace(strName, ":", "")
    strName = Replace(strName,"/","")
    strName = Replace(strName,"\","")
    strName = Replace(strName,",","")
    strName = Replace(strName, Chr(34),"")
    strName = Replace(strName,Chr(39),"")
    strName = Replace(strName,"?","")

    strName = "C:\Test\" & strName & ".msg"
    objMessage.SaveAs strName, olMSG  
Next
  


