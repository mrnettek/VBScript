Const olSentMail = 5

Set objDictionary = CreateObject("Scripting.Dictionary")
Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olSentMail)

Set colItems = objFolder.Items

For Each objItem in colItems
    Set colRecipients = objItem.Recipients
    For Each objRecipient in colRecipients
        strAddress = objRecipient.Address
        If Not objDictionary.Exists(strAddress) Then
            objDictionary.Add strAddress, strAddress
        End If
    Next
Next

i = 1

For Each strKey in objDictionary.Keys
    objWorksheet.Cells(i, 1) = strKey
    i = i + 1
Next
  


