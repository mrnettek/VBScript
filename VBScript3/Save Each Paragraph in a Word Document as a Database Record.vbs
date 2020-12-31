Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = C:\Scripts\Test.mdb" 

objRecordSet.Open "SELECT * FROM WordParagraphs" , _
    objConnection, adOpenStatic, adLockOptimistic

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

Set objSelection = objWord.Selection

For i = 1 to objDoc.Paragraphs.Count
    objDoc.Paragraphs(i).Range.Select
    If Len(objSelection.Text) > 1 Then
        objRecordSet.AddNew
        objRecordSet("ParagraphNumber") = i
        objRecordSet("ParagraphText") = objSelection.Text
        objRecordSet.Update
    End If
Next
  


