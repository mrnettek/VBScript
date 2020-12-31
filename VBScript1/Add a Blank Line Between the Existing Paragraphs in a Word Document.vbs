Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

Set colParagraphs = objDoc.Paragraphs
colParagraphs.LineUnitAfter = 1
  


