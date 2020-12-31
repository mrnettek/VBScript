Set objDictionary = CreateObject("Scripting.Dictionary")

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Sample.doc")

Set colWords = objDoc.Words

For Each strWord in colWords
    strWord = LCase(strWord)
    If objDictionary.Exists(strWord) Then
    Else
        objDictionary.Add strWord, strWord
   End If
Next

Set objDoc2 = objWord.Documents.Add()
Set objSelection = objWord.Selection

For Each strItem in objDictionary.Items
    objSelection.TypeText strItem & vbCrLf
Next

Set objRange = objDoc2.Range
objRange.Sort
  


