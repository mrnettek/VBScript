Const wdTitleWord = 2
Const wdTitleSentence = 4

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

Set colParagraphs = objDoc.Paragraphs

For Each objParagraph in colParagraphs
    If objParagraph.Range.Font.Size = 14 Then
        objParagraph.Range.Case = wdTitleWord
    Else
        objParagraph.Range.Case = wdTitleSentence
    End If
Next
  


